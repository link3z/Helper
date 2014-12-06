Imports System
Imports System.Net
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports System.Collections

Namespace Web
    ''' <summary>
    ''' Clase que carga una página HTML y permite acceder a todos sus bloques y propiedades
    ''' de forma cómoda
    ''' </summary>
    Public Class cPaginaHTML
#Region " DECLARACIONES "                                                                                
        ''' <summary>
        ''' Timeout para la carga de la página
        ''' </summary>
        Private iTimeOut As Integer = 30
#End Region

#Region " CONSTRUCTORES "
        ''' <summary>
        ''' Crea una nueva instancia de la clase sin cargar ninguna págiga
        ''' </summary>
        Public Sub New(Optional ByVal eTimeOut As Integer = 30)
            iTimeOut = eTimeOut
        End Sub

        Public Sub Dispose()
            ' Si se estaba cargando la página se detiene la carga antes de destruir
            ' el objeto para evitar problemas
            Me.Cancel = True
        End Sub
#End Region

#Region " PROPIEDADES "
        ''' <summary>
        ''' Permite cancelar la carga de la página
        ''' </summary>
        Public Property Cancel() As Boolean
            Get
                Return iCancel
            End Get
            Set(ByVal value As Boolean)
                iCancel = value
            End Set
        End Property
        Private iCancel As Boolean = False

        ''' <summary>
        ''' URL de la página cargada
        ''' </summary>
        Public ReadOnly Property URL() As String
            Get
                Return iURL
            End Get
        End Property
        Private iURL As String = ""

        ''' <summary>
        ''' Host de la página
        ''' </summary>
        Public ReadOnly Property Host() As String
            Get
                Return iHost
            End Get
        End Property
        Private iHost As String = ""

        ''' <summary>
        ''' URL del servidor
        ''' </summary>
        Public ReadOnly Property ServerURL() As String
            Get
                Return iServerURL
            End Get
        End Property
        Private iServerURL As String = ""

        ''' <summary>
        ''' Path de la URL
        ''' </summary>
        Public ReadOnly Property PathURL() As String
            Get
                Return iPathURL
            End Get
        End Property
        Private iPathURL As String = ""

        ''' <summary>
        ''' Código fuente de la página
        ''' </summary>
        Public Property Source() As String
            Get
                Return iSource
            End Get
            Set(ByVal value As String)
                iSource = value
                iHost = ""
                iCharacterSet = ""
                iContentEncoding = ""
                iContentLength = 0
                iContentType = ""
                iLastModified = ""
            End Set
        End Property
        Private iSource As String = ""

        ''' <summary>
        ''' CharacterSet utilizado en la página
        ''' </summary>
        Public ReadOnly Property CharacterSet() As String
            Get
                Return iCharacterSet
            End Get
        End Property
        Private iCharacterSet As String = ""

        ''' <summary>
        ''' ContentEnconding utilizado en la página
        ''' </summary>
        Public ReadOnly Property ContentEncoding() As String
            Get
                Return iContentEncoding
            End Get
        End Property
        Private iContentEncoding As String = ""

        ''' <summary>
        ''' Tamaño del contenido de la página
        ''' </summary>
        Public ReadOnly Property ContentLength() As Long
            Get
                Return iContentLength
            End Get
        End Property
        Private iContentLength As Long = 0

        ''' <summary>
        ''' Tipo del contenido de la página
        ''' </summary>
        Public ReadOnly Property ContentType() As String
            Get
                Return (iContentType)
            End Get
        End Property
        Private iContentType As String = ""

        ''' <summary>
        ''' Fecha de la última modificación de la página
        ''' </summary>
        Public ReadOnly Property LastModified() As String
            Get
                Return (iLastModified)
            End Get
        End Property
        Private iLastModified As String = ""

        ''' <summary>
        ''' Cabecera de la página
        ''' </summary>
        Public ReadOnly Property Head() As String
            Get
                Return (obtenerContenidoTag("Head", iSource))
            End Get
        End Property

        ''' <summary>
        ''' Title de la página
        ''' </summary>
        Public ReadOnly Property Title() As String
            Get
                ' Se eliminan todos los comentarios
                Dim paraDevolver As String = eliminarComentarios(iSource)

                ' Se obtiene el título de la página
                paraDevolver = obtenerContenidoTag("Title", paraDevolver)

                ' Se eliminan los TAGS del título
                Dim strPattern As String = "<[^>]*>"
                paraDevolver = Regex.Replace(paraDevolver, strPattern, "")

                Return (paraDevolver.Trim())
            End Get
        End Property

        ''' <summary>
        ''' Body de la página
        ''' </summary>
        Public ReadOnly Property Body() As String
            Get
                ' Si no se pudiera encontrar la etiqueta Body se devuelve todo
                ' el código fuente de la página ya que se supone que todo es la propia página
                Dim strBody As String = obtenerContenidoTag("Body", iSource)
                If strBody = "" Then
                    strBody = Me.Source
                End If
                Return (strBody)
            End Get
        End Property

        ''' <summary>
        ''' Obtiene unicamente el texto de la página, eliminando los scripts
        ''' y los tags y convirtiendo todos los caracteres especiales
        ''' </summary>
        Public ReadOnly Property Text() As String
            Get
                Dim r As New Regex("")
                Dim opts As RegexOptions = RegexOptions.IgnoreCase Or RegexOptions.Singleline

                ' Se obtiene el cuerpo de la página
                Dim strText As String = Me.Body

                ' Se eliminan todos los comentarios
                strText = eliminarComentarios(strText)

                ' Se eliminan todos los scripts
                Dim strPattern As String = obtenerContenidoTag("SCRIPT")
                strText = Regex.Replace(strText, strPattern, "", opts)

                ' Se eliminan todos los tags
                strPattern = "<[^>]*>"
                strText = Regex.Replace(strText, strPattern, " ", opts)

                ' Se convierten todos los caracteres a ASCII
                strText = HTML.ISO2ASCII(strText)
                strText = Regex.Replace(strText, "&", "&", opts)

                ' Se eliminan todos los espacios dobles que puderan quedar
                Dim m As System.Text.RegularExpressions.MatchCollection
                Do
                    strText = Regex.Replace(strText, "\s\s", " ")
                    m = Regex.Matches(strText, "\s\s")
                Loop While m.Count > 0

                Return (strText.Trim())
            End Get
        End Property
#End Region

#Region " METODOS "
        ''' <summary>
        ''' Carga una página y rellena todas las propiedades para poder utilizar el objeto
        ''' </summary>
        ''' <param name="eURL">URL que se va a cargar</param>
        ''' <returns>True o False dependiendo de si se pudo ejecutar la operación correctamente</returns>
        Public Function cargarURL(ByVal eURL As String) As Boolean            
            iURL = eURL

            ' Variables para la carga de la página
            Const DEFAULT_CONTENT_LENGTH As Integer = 40000
            Dim strSource As String = ""
            Dim strHost As String = ""
            Dim strServerURL As String = ""
            Dim strPathURL As String = ""
            Dim strCharacterSet As String = ""
            Dim strContentEncoding As String = ""
            Dim lngContentLength As Long = 0
            Dim strContentType As String = ""
            Dim strLastModified As String = ""
            Dim intTotalLength As Integer = 0

            ' Se comprueba que se está cargando una página, si la dirección no 
            ' está establecida no se puede cargar
            If (iURL Is Nothing) OrElse (iURL = "") Then
                If Log._LOG_ACTIVO Then Log.escribirLog("La dirección de la página no puede ser una cadena vacía.", , New StackTrace(0, True))
                Return False
            Else
                Try
                    ' Se solicita la página al servidor
                    Dim hrqURL As HttpWebRequest = DirectCast(HttpWebRequest.Create(iURL), HttpWebRequest)
                    Dim hrspURL As HttpWebResponse = DirectCast(hrqURL.GetResponse(), HttpWebResponse)
                    Dim srdrInput As New StreamReader(hrspURL.GetResponseStream())
                    Dim chrBuff As Char() = New Char(255) {}
                    Dim intLen As Integer

                    ' Se obtiene el tamaño de la página
                    If lngContentLength <= 0 Then
                        lngContentLength = DEFAULT_CONTENT_LENGTH
                    End If

                    ' Se fija el TimeOut de carga
                    Dim tmeExpire As New DateTime(DateTime.Now.Ticks)
                    tmeExpire = tmeExpire.AddSeconds(iTimeOut)

                    ' Se realiza un bucle hasta que se carga toda la página o se cancela la carga
                    Do
                        If iCancel Then
                            If Log._LOG_ACTIVO Then Log.escribirLog("Se canceló la carga de la página '" & iURL & "'.", , New StackTrace(0, True))
                            Return False
                        End If

                        intLen = srdrInput.Read(chrBuff, 0, 256)
                        Dim strBuff As New String(chrBuff, 0, intLen)
                        strSource = strSource + strBuff
                        intTotalLength = intTotalLength + intLen
                        If intTotalLength > lngContentLength Then
                            lngContentLength = 2 * intTotalLength
                        End If

                        ' Se verifica que no se superara el TimeOut establecido
                        If System.DateTime.Compare(tmeExpire, System.DateTime.Now) < 0 Then
                            If Log._LOG_ACTIVO Then Log.escribirLog("Se superó el TimeOut (" & iTimeOut & ") al cargar '" & iURL & "'.", , New StackTrace(0, True))
                            Return False
                        End If
                    Loop While (intLen > 0)

                    srdrInput.Close()
                    hrspURL.Close()

                    ' Se guardan todas las viriables e información de la página
                    strHost = hrspURL.ResponseUri.Host

                    Dim m As Match = Regex.Match(hrspURL.ResponseUri.AbsoluteUri, "/", RegexOptions.RightToLeft)
                    If m Is Nothing Then
                        strPathURL = hrspURL.ResponseUri.AbsoluteUri & "/"
                    Else
                        strPathURL = hrspURL.ResponseUri.AbsoluteUri.Substring(0, m.Index) & "/"
                    End If

                    m = Regex.Match(hrspURL.ResponseUri.AbsoluteUri, strHost, RegexOptions.RightToLeft Or RegexOptions.IgnoreCase)
                    If m Is Nothing Then
                        strServerURL = hrspURL.ResponseUri.AbsoluteUri
                    Else
                        strServerURL = hrspURL.ResponseUri.AbsoluteUri.Substring(0, m.Index + strHost.Length)
                    End If

                    strCharacterSet = hrspURL.CharacterSet
                    strContentEncoding = hrspURL.ContentEncoding
                    lngContentLength = hrspURL.ContentLength
                    strContentType = hrspURL.ContentType
                    strLastModified = hrspURL.LastModified.ToString()
                Catch ex As Exception
                    If Log._LOG_ACTIVO Then Log.escribirLog("ERROR al cargar la página '" & iURL & "'.", ex, New StackTrace(0, True))
                    Return False
                End Try
            End If

            ' Se guardan todos los datos obtenidos en las propiedades del objeto
            iHost = strHost
            iServerURL = strServerURL
            iPathURL = strPathURL
            iSource = strSource
            iCharacterSet = strCharacterSet
            iContentEncoding = strContentEncoding
            iContentLength = lngContentLength
            iContentType = strContentType
            iLastModified = strLastModified

            ' Si el código llega hasta este punto todo está correcto
            Return True
        End Function

        ''' <summary>
        ''' Obtienen todos los tags de un determinado tipo
        ''' </summary>
        ''' <param name="eTipo">Tipo de los tags que se quieren obtener</param>
        ''' <returns>Listado con todos los tags de un determinado tipo</returns>
        Public Function obtenerTagsPorTipo(ByVal eTipo As String) As String()
            Return (obtenerTagsPorTipo(eTipo, iSource))
        End Function

        ''' <summary>
        ''' Obtiene todos los tags de un determinado tipo
        ''' </summary>
        ''' <param name="eTipo">Tipo de los tags que se quieren obtener</param>
        ''' <param name="eSource">Código o bloque donde se quieren obtener los tags</param>
        ''' <returns>Listado con todos los tags de un determinado tipo</returns>
        Public Function obtenerTagsPorTipo(ByVal eTipo As String, _
                                           ByVal eSource As String) As String()
            Dim opts As RegexOptions = RegexOptions.IgnoreCase Or RegexOptions.Singleline
            Dim strPattern As String

            ' Se eliminan los comentarios
            If eTipo <> "!" Then
                eSource = eliminarComentarios(eSource)
                strPattern = "<(?<TagName>" & eTipo & ")(>|\s+[^>]*>)"
            Else
                strPattern = "<(?<TagName>" & eTipo & ")--"
            End If

            ' Se busca la primera posición
            Dim mc As MatchCollection = Regex.Matches(eSource, strPattern, opts)

            ' Se carga el contenido de cada tag
            Dim strTagContents As New ArrayList()
            For Each m As Match In mc
                strTagContents.Add(obtenerContenidoTag(eTipo, eSource.Substring(m.Groups("TagName").Index - 1)))
            Next
            Return DirectCast(strTagContents.ToArray(GetType([String])), String())
        End Function

        ''' <summary>
        ''' Obtiene todos los HREFS de la página
        ''' </summary>
        ''' <returns>Listado con todos los HREFS de la página</returns>
        Public Function obtenerHREFS() As String()
            ' Elimina todos los comentarios
            Dim strSource As String = eliminarComentarios(iSource)

            ' Se buscan todos los HREFS mediante la expresión regular
            Dim r As New Regex("<a[^>]*href\s*=\s*""?(?<HRef>[^"">\s]*)""?[^>]*>", RegexOptions.IgnoreCase Or RegexOptions.Singleline)
            Dim mc As MatchCollection = r.Matches(Source)
            Dim strHRefs As New ArrayList()

            ' Se recorrent odos los resultados para ir limiando los resultados
            For Each m As Match In mc
                Dim strHRef As String = m.Groups("HRef").Value
                strHRef.Trim()

                ' Se normaliza la URL
                If strHRef <> "" Then
                    If Left(strHRef, 1) <> "#" Then
                        If Left(strHRef, 1) = "/" Then
                            strHRef = iServerURL + strHRef
                        Else
                            If [String].Compare(Left(strHRef, 7), "http://", True) <> 0 Then
                                If [String].Compare(Left(strHRef, 3), "www", True) = 0 Then
                                    strHRef = "http://" & strHRef
                                Else
                                    strHRef = iPathURL + strHRef
                                End If
                            End If
                        End If
                    End If
                End If
                strHRefs.Add(strHRef)
            Next

            ' Se devuelve la lista
            Return DirectCast(strHRefs.ToArray(GetType(String)), String())
        End Function

        ''' <summary>
        ''' Obtienen todos los tags de la página
        ''' </summary>
        ''' <returns>Listado con todos los TAGS de la página</returns>
        ''' <remarks></remarks>
        Public Function obtenerTags() As String()
            Dim strPatternTag As String = "<(?<Comment>!?)(?<Tag>[^>/\s]+)(>|\s+[^>]*>)"
            Dim mc As MatchCollection = Regex.Matches(iSource, strPatternTag, RegexOptions.IgnoreCase Or RegexOptions.Singleline)
            Dim intLength As Integer = iSource.Length
            Dim i As Integer
            Dim strTagList As New ArrayList()
            Dim strTagName As String
            Dim strTagData As String
            For i = 0 To mc.Count - 1
                If iCancel Then
                    Return Nothing
                End If

                ' Se obtiene el nombre del tag capturado
                If mc(i).Groups("Comment").Value = "" Then
                    strTagName = mc(i).Groups("Tag").Value
                Else
                    strTagName = "!"
                End If

                ' Se busca el la posición de finalización del tag 
                strTagData = obtenerContenidoTag(strTagName, iSource.Substring(mc(i).Groups("Tag").Index - 1))
                strTagList.Add(strTagData)
            Next

            Return DirectCast(strTagList.ToArray(GetType(String)), String())
        End Function


        ''' <summary>
        ''' Obtiene el contenido de un determinado tag
        ''' </summary>
        ''' <param name="eTag">Nombre del tag</param>
        ''' <returns>Contenido del tag</returns>
        ''' <remarks></remarks>
        Private Function obtenerContenidoTag(ByVal eTag As String) As String
            Dim strPatternTag As String
            If eTag = "!" Then
                strPatternTag = "<!.*?-->"
            ElseIf String.Compare(eTag, "!doctype", True) = 0 Then
                strPatternTag = "<!doctype.*?>"
            ElseIf [String].Compare(eTag, "br", True) = 0 Then
                strPatternTag = "<br\s*/?\s*>"
            Else
                strPatternTag = "<(" & eTag & ")(>|\s+[^>]*>).*?</\1\s*>"
            End If
            Return (strPatternTag)
        End Function

        ''' <summary>
        ''' Obtienen todo el contenido de un tag hasta que encuentra su cierre o el inicio de otro tag
        ''' </summary>
        ''' <param name="eTag">Tag que se va a comprobar</param>
        ''' <param name="eSource">Donde se va a comprobar el tag</param>
        ''' <returns>Contenido del tag</returns>
        Private Function obtenerContenidoTag(ByVal eTag As String, ByVal eSource As String) As String
            Dim strPatternTag As String = obtenerContenidoTag(eTag)
            Dim strPatternTagNoClose As String = "<" & eTag & "(>|\s+[^>]*>)[^<]"
            Dim opts As RegexOptions = RegexOptions.IgnoreCase Or RegexOptions.Singleline
            Dim m As Match
            Dim strGetTagByName As String

            m = System.Text.RegularExpressions.Regex.Match(eSource, strPatternTag, opts)
            If m.Value = "" Then
                m = System.Text.RegularExpressions.Regex.Match(eSource, strPatternTagNoClose, opts)
                If m Is Nothing Then
                    strGetTagByName = eSource
                Else
                    strGetTagByName = m.Value
                End If
            Else
                strGetTagByName = m.Value
            End If
            Return (strGetTagByName)
        End Function

        ''' <summary>
        ''' Elimina los comentarios el código HTML
        ''' </summary>
        ''' <param name="eSource">Donde se tiene que eliminar los comentarios</param>
        ''' <returns>Código original sin los comentarios</returns>
        Private Function eliminarComentarios(ByVal eSource As String) As String
            Dim r As New Regex(obtenerContenidoTag("!"))
            Return (r.Replace(eSource, ""))
        End Function


#End Region

    End Class
End Namespace

