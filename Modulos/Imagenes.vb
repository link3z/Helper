Imports System.Runtime.InteropServices
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms
Imports System.Drawing.Printing
Imports System.Net

Namespace Imagenes
    Public Module modImagenes
#Region " PROPIEDADES "
        ''' <summary>
        ''' Listado con todos los codecs para el tratamiento de imagen disponibles
        ''' en el equipo
        ''' </summary>
        Public ReadOnly Property codecsImagen() As Dictionary(Of String, ImageCodecInfo)
            Get
                If iCodecsImagen Is Nothing Then
                    iCodecsImagen = New Dictionary(Of String, ImageCodecInfo)()

                    For Each codec As ImageCodecInfo In ImageCodecInfo.GetImageEncoders()
                        iCodecsImagen.Add(codec.MimeType.ToLower(), codec)
                    Next
                End If
                Return iCodecsImagen
            End Get
        End Property
        Private iCodecsImagen As Dictionary(Of String, ImageCodecInfo) = Nothing
#End Region

#Region " CAMBIO TAMAÑO, RESOLUCIÓN Y TRANSFORMACIONES "
        ''' <summary>
        ''' Cambia el tamaño a una imagen que se le pasa como parámetro
        ''' </summary>
        ''' <param name="eImagen">Imagen a la que se le va a cambiar el tamaño</param>
        ''' <param name="eAncho">Ancho de la imagen resultante</param>
        ''' <param name="eAlto">Alto de la imagen resultante</param>
        ''' <returns>Bitmap con la nueva imagen aplicando el camibo de resolución</returns>
        Public Function resizeImage(ByVal eImagen As System.Drawing.Image, _
                                    ByVal eAncho As Integer, _
                                    ByVal eAlto As Integer) As System.Drawing.Bitmap
            If eImagen Is Nothing Then Return eImagen

            ' Se crea la nueva imagen con los nuevos tamaños indicados y la resolución
            ' de la imagen original
            Dim paraDevolver As Bitmap = Nothing
            Try
                paraDevolver = New Bitmap(eAncho, eAlto)
                paraDevolver.SetResolution(eImagen.HorizontalResolution, eImagen.VerticalResolution)

                Using elGraphics As Graphics = Graphics.FromImage(paraDevolver)
                    elGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
                    elGraphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic
                    elGraphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality
                    elGraphics.DrawImage(eImagen, 0, 0, paraDevolver.Width, paraDevolver.Height)
                End Using
            Catch ex As Exception
                If Log._LOG_ACTIVO Then Log.escribirLog("Se ha producido un error al cambiar la resolución de la imagen...", ex, New StackTrace(0, True))
                paraDevolver = Nothing
            End Try

            Return paraDevolver
        End Function

        ''' <summary>
        ''' Realiza un ajuste de tamaño de una imagen manteniendo la proporción de la imagen original
        ''' </summary>
        ''' <param name="eImagen">Imagen que se va a escalar</param>
        ''' <param name="eAncho">Ancho máximo que puede tener la imagen resultante</param>
        ''' <param name="eAlto">Alto máximo que puede tener la imagen resultante</param>
        ''' <returns>Imagen escalada manteniendo la relación de aspecto original</returns>
        Public Function ajustarTamanho(ByVal eImagen As Image, _
                                       ByVal eAncho As Integer, _
                                       ByVal eAlto As Integer) As Image
            If eImagen Is Nothing Then Return Nothing

            Try
                Dim AnchoOriginal As Integer = eImagen.Width
                Dim AltoOriginal As Integer = eImagen.Height
                Dim XOriginal, YOriginal, XDestino, YDestino As Integer
                Dim Porcentaje, PorcentajeAncho, PorcentajeAlto As Single

                ' Se calculan las relaciones de aspecto
                PorcentajeAncho = (CSng(eAncho) / CSng(AnchoOriginal))
                PorcentajeAlto = (CSng(eAlto) / CSng(AltoOriginal))

                ' Se selecciona la propiorción de aspecto a utilizar
                If PorcentajeAlto < PorcentajeAncho Then
                    Porcentaje = PorcentajeAlto
                    XDestino = System.Convert.ToInt16((eAncho - (AnchoOriginal * Porcentaje)) / 2)
                Else
                    Porcentaje = PorcentajeAncho
                    YDestino = System.Convert.ToInt16((eAlto - (AltoOriginal * Porcentaje)) / 2)
                End If

                Dim AnchoDestino As Integer = CInt(Math.Truncate(AnchoOriginal * Porcentaje))
                Dim AltoDestino As Integer = CInt(Math.Truncate(AltoOriginal * Porcentaje))

                ' Se crea la nueva imagen con las nuevas proporciones
                Dim bmPhoto As New Bitmap(AnchoDestino, AltoDestino)
                bmPhoto.SetResolution(eImagen.HorizontalResolution, eImagen.VerticalResolution)

                Dim grPhoto As Graphics = Graphics.FromImage(bmPhoto)
                grPhoto.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic

                grPhoto.DrawImage(eImagen, New Rectangle(0, 0, AnchoDestino, AltoDestino), New Rectangle(XOriginal, YOriginal, AnchoOriginal, AltoOriginal), GraphicsUnit.Pixel)

                grPhoto.Dispose()
                Return bmPhoto
            Catch ex As Exception
                ' Si se produce eun error se devuelve la imagen original
                If Log._LOG_ACTIVO Then Log.escribirLog("Se ha producido un error al tratar de realizar un escalado de la imagen...", ex, New StackTrace(0, True))
                Return eImagen
            End Try
        End Function

        ''' <summary>
        ''' Rota la imagen en el punto especificado los grados indicados
        ''' </summary>
        ''' <param name="Imagen">Imagen que se quiere rotar</param>
        ''' <param name="Angulo">Ángulo de rotación</param>
        ''' <returns>Imagen rotada</returns>
        Public Function rotarImagen(ByVal Imagen As Image, _
                                    ByVal Angulo As Single) As Bitmap
            Dim mRotate As New Matrix()
            mRotate.Translate(Imagen.Width / -2, Imagen.Height / -2, MatrixOrder.Append)
            mRotate.RotateAt(Angulo, New System.Drawing.Point(0, 0), MatrixOrder.Append)
            Using gp As New GraphicsPath()
                gp.AddPolygon(New System.Drawing.Point() {New System.Drawing.Point(0, 0), New System.Drawing.Point(Imagen.Width, 0), New System.Drawing.Point(0, Imagen.Height)})
                gp.Transform(mRotate)
                Dim pts As System.Drawing.PointF() = gp.PathPoints

                Dim bbox As Rectangle = boundingBox(Imagen, mRotate)
                Dim bmpDest As New Bitmap(bbox.Width, bbox.Height)

                Using gDest As Graphics = Graphics.FromImage(bmpDest)
                    Dim mDest As New Matrix()
                    mDest.Translate(bmpDest.Width / 2, bmpDest.Height / 2, MatrixOrder.Append)
                    gDest.Transform = mDest
                    gDest.DrawImage(Imagen, pts)
                    Return bmpDest
                End Using
            End Using
        End Function

        Private Function boundingBox(ByVal img As Image, ByVal matrix As Matrix) As Rectangle
            Dim gu As New GraphicsUnit()
            Dim rImg As Rectangle = Rectangle.Round(img.GetBounds(gu))

            ' Transform the four points of the image, to get the resized bounding box.
            Dim topLeft As New System.Drawing.Point(rImg.Left, rImg.Top)
            Dim topRight As New System.Drawing.Point(rImg.Right, rImg.Top)
            Dim bottomRight As New System.Drawing.Point(rImg.Right, rImg.Bottom)
            Dim bottomLeft As New System.Drawing.Point(rImg.Left, rImg.Bottom)
            Dim points As System.Drawing.Point() = New System.Drawing.Point() {topLeft, topRight, bottomRight, bottomLeft}
            Dim gp As New GraphicsPath(points, New Byte() {CByte(PathPointType.Start), CByte(PathPointType.Line), CByte(PathPointType.Line), CByte(PathPointType.Line)})
            gp.Transform(matrix)
            Return Rectangle.Round(gp.GetBounds())
        End Function
#End Region

#Region " CODECS Y MIME TYPES "
        ''' <summary> 
        ''' Obtiene el codec asociado a un tipo myme
        ''' </summary> 
        ''' <param name="eMime">Tipo mime del que se quiere obtener información del code</param>
        Public Function obtenerCodec(ByVal eMime As String) As ImageCodecInfo
            Dim paraDevolver As ImageCodecInfo = Nothing

            Dim elNombre As String = eMime.ToLower()
            If codecsImagen.ContainsKey(elNombre) Then
                paraDevolver = codecsImagen(elNombre)
            End If

            Return paraDevolver
        End Function
#End Region

#Region " EXPORTACIÓN DE IMÁGENES "
        ''' <summary>
        ''' Guarda una imagen en formato JPG utilizando los codecs
        ''' Mime y los parámetros que se le pasan como parámetros
        ''' </summary>
        ''' <param name="eRutaSalida">Ruta de salida de la imagen</param>
        ''' <param name="eImagen">Imagen que se quiere guardar</param>
        ''' <param name="eCalidad">Calidad a aplicar a la imagen</param>
        Public Sub GuardarComoJPEG(ByVal eRutaSalida As String, _
                                   ByVal eImagen As Image, _
                                   Optional ByVal eCalidad As Integer = 90)
            ' Si la calidad de la imagen no se encuentra entre los parámetros
            ' validos asignamos el máximo de calidad
            If (eCalidad < 0) OrElse (eCalidad > 100) Then
                eCalidad = 90
            End If

            ' Se crean los parámetros asociados al codec
            Dim laCalidad As New EncoderParameter(System.Drawing.Imaging.Encoder.Quality, eCalidad)
            Dim elCodec As ImageCodecInfo = obtenerCodec("image/jpeg")

            ' Se guarda l aimagen utilizando los parámetros generados
            Dim encoderParams As New EncoderParameters(1)
            encoderParams.Param(0) = laCalidad
            eImagen.Save(eRutaSalida, elCodec, encoderParams)
        End Sub
#End Region

#Region " IMPORTACIÓN DE IMAGENES "
        ''' <summary>
        ''' Obtine una imagen desde una URL especificada
        ''' </summary>
        ''' <param name="eURL">Dirección URL de la imagen</param>
        ''' <returns>Imagen obtenida de la dirección URL</returns>
        Public Function obtenerImagenHTTP2Image(ByVal eURL As String) As Image
            Dim paraDevolver As Image = Nothing

            Dim elRequest As Net.HttpWebRequest = Nothing
            Dim elResponse As Net.HttpWebResponse = Nothing
            Dim elStream As IO.Stream = Nothing

            Try
                elRequest = Net.WebRequest.Create(eURL)
                elRequest.AllowWriteStreamBuffering = True

                elResponse = elRequest.GetResponse()
                elStream = elResponse.GetResponseStream()

                If Not elStream Is Nothing Then paraDevolver = Image.FromStream(elStream)
            Catch ex As System.Exception
                If Log._LOG_ACTIVO Then Log.escribirLog("Se ha producido un error al obtener una imagen desde una URL...", ex, New StackTrace(0, True))
                paraDevolver = Nothing
            Finally
                If Not elStream Is Nothing Then elStream.Close()
                If Not elResponse Is Nothing Then elResponse.Close()
            End Try

            Return paraDevolver
        End Function

        ''' <summary>
        ''' Obtine una imagen desde una URL especificada devolviendo el array de bytes que la representa
        ''' </summary>
        ''' <param name="eURL">Dirección URL de la imagen</param>
        ''' <returns>Imagen obtenida de la dirección URL</returns>
        Public Function obtenerImagenHTTP2Byte(ByVal eURL As String) As Byte()
            Dim paraDevolver As Byte() = Nothing

            Dim elRequest As Net.HttpWebRequest = Nothing
            Dim elResponse As Net.HttpWebResponse = Nothing

            Try
                elRequest = Net.HttpWebRequest.Create(eURL)
                elResponse = elRequest.GetResponse

                Using elStream = elResponse.GetResponseStream()
                    Using ms As New MemoryStream(elResponse.ContentLength - 1)
                        Dim bytesLeidos As Long = 0
                        Dim buffer As Byte() = New Byte(256) {}
                        Do
                            bytesLeidos = elStream.Read(buffer, 0, buffer.Length)
                            ms.Write(buffer, 0, bytesLeidos)
                        Loop While bytesLeidos > 0

                        paraDevolver = ms.ToArray
                    End Using
                End Using

                Return paraDevolver
            Catch ex As System.Exception
                If Log._LOG_ACTIVO Then Log.escribirLog("Se ha producido un error al obtener una imagen desde una URL...", ex, New StackTrace(0, True))
                paraDevolver = Nothing
            End Try

            Return paraDevolver
        End Function
#End Region

#Region " MARCA DE AGUA "
        ''' <summary>
        ''' Añade una marca de agua a una imagen en unas coordenadas especificadas
        ''' </summary>
        ''' <param name="eMarcaAgua">Imagen a utilizar como marca de agua</param>
        ''' <param name="eImgen">Imagen donde se insertará la marca de agua</param>
        ''' <param name="x">Coordenada X</param>
        ''' <param name="y">Coordenada Y</param>
        ''' <returns>True o False dependiendo de si se pudo ejecutar correctamente</returns>
        Public Function anhadirMarcaAgua(ByVal eMarcaAgua As Bitmap, _
                                         ByVal eImgen As Bitmap, _
                                         ByVal x As Integer, _
                                         ByVal y As Integer) As Boolean
            Try
                ' Utilizado para la generación de la transparencia
                Dim ALPHA As Byte = 128
                Dim clr As Color

                For py As Integer = 0 To eMarcaAgua.Height - 1
                    For px As Integer = 0 To eMarcaAgua.Width - 1
                        clr = eMarcaAgua.GetPixel(px, py)
                        eMarcaAgua.SetPixel(px, py, Color.FromArgb(ALPHA, clr.R, clr.G, clr.B))
                    Next px
                Next py

                eMarcaAgua.MakeTransparent(eMarcaAgua.GetPixel(0, 0))

                Dim gr As Graphics = Graphics.FromImage(eImgen)
                gr.DrawImage(eMarcaAgua, x, y)
            Catch ex As Exception
                If Log._LOG_ACTIVO Then Log.escribirLog("Se ha producido un error al añadir la marca de agua...", ex, New StackTrace(0, True))
                Return False
            End Try

            Return True
        End Function
#End Region

#Region " CONVERSIONES "
        ''' <summary>
        ''' Transforma cualquier imagen en un icono.
        ''' </summary>
        ''' <param name="Imagen">Imagen a partir de la cual vamos a generar el icono</param>
        ''' <returns>El icono generado</returns>
        Public Function imagenAIcono(ByVal Imagen As Image) As Icon
            Dim bm As New Bitmap(Imagen, 16, 16)
            Return System.Drawing.Icon.FromHandle(bm.GetHicon)
        End Function

        ''' <summary>
        ''' Transforma cualquier icono en imagen.
        ''' </summary>
        ''' <param name="Icono">icono a partir del que vamos a generar la imagen</param>
        ''' <returns>La imagen generada</returns>
        Public Function iconoAImagen(ByVal Icono As Icon) As Image
            Dim Imagen As Image = Icono.ToBitmap
            Return Imagen
        End Function

        ''' <summary>
        ''' Obtiene un cursor para el ratón a partir de una imagen
        ''' </summary>
        ''' <param name="eImagen">Imagen que se va a utilizar como cursor</param>
        ''' <returns>Cursor par ausar en el ratón</returns>
        Public Function imagenACursor(ByVal eImagen As Image) As Cursor
            Try
                Dim ptr As IntPtr = DirectCast(eImagen, Bitmap).GetHicon()
                Return New Cursor(ptr)
            Catch ex As Exception
                Return Cursors.Default
            End Try
        End Function
#End Region

#Region " COLORES "
        ''' <summary>
        ''' Reemplaza un color de una imagen por otro aplicando una tolerancia al cambio
        ''' </summary>
        ''' <param name="eImagen">Imagen original</param>
        ''' <param name="eColorAntiguo">Color antiguo</param>
        ''' <param name="eColorNuevo">Color nuevo</param>
        ''' <param name="eTolerancia">Tolerancia</param>
        ''' <returns>La imagen con el color cambiado</returns>
        Public Function reemplazarColor(ByVal eImagen As Image, _
                                        ByVal eColorAntiguo As Color, _
                                        ByVal eColorNuevo As Color, _
                                        ByVal eTolerancia As Integer) As Image
            If eImagen Is Nothing Then Return Nothing

            Dim elBitmap As Bitmap = DirectCast(eImagen.Clone(), Bitmap)

            Dim c As Color
            Dim iR_Min As Integer, iR_Max As Integer
            Dim iG_Min As Integer, iG_Max As Integer
            Dim iB_Min As Integer, iB_Max As Integer

            ' Se calcula la tolerancia en RGB
            ' R
            iR_Min = Math.Max(CInt(eColorAntiguo.R) - eTolerancia, 0)
            iR_Max = Math.Min(CInt(eColorAntiguo.R) + eTolerancia, 255)

            ' G
            iG_Min = Math.Max(CInt(eColorAntiguo.G) - eTolerancia, 0)
            iG_Max = Math.Min(CInt(eColorAntiguo.G) + eTolerancia, 255)

            ' B
            iB_Min = Math.Max(CInt(eColorAntiguo.B) - eTolerancia, 0)
            iB_Max = Math.Min(CInt(eColorAntiguo.B) + eTolerancia, 255)


            ' Se recorre la matriz de colores del bitmap pixel a pixel y se cambia
            ' el color si está dentro de la tolerancia configurada
            For x As Integer = 0 To elBitmap.Width - 1
                For y As Integer = 0 To elBitmap.Height - 1
                    c = elBitmap.GetPixel(x, y)

                    If (c.R >= iR_Min AndAlso c.R <= iR_Max) AndAlso (c.G >= iG_Min AndAlso c.G <= iG_Max) AndAlso (c.B >= iB_Min AndAlso c.B <= iB_Max) Then
                        If eColorNuevo = Color.Transparent Then
                            elBitmap.SetPixel(x, y, Color.FromArgb(0, eColorNuevo.R, eColorNuevo.G, eColorNuevo.B))
                        Else
                            elBitmap.SetPixel(x, y, Color.FromArgb(c.A, eColorNuevo.R, eColorNuevo.G, eColorNuevo.B))
                        End If
                    End If
                Next
            Next

            Return DirectCast(elBitmap.Clone(), Image)
        End Function

        ''' <summary>
        ''' Convierte una imagen a escala de grises
        ''' </summary>
        ''' <param name="eImagen">Imagen original</param>
        ''' <param name="eAncho">Ancho resultante</param>
        ''' <param name="eAlto">Alto resultante</param>
        ''' <returns>Imagen convertida a escala de grises y con el tamaño especifiado</returns>
        Public Function obtenerEscalaGrises(ByVal eImagen As Bitmap, _
                                            ByVal eAncho As Integer, _
                                            ByVal eAlto As Integer) As Bitmap
            If eImagen Is Nothing Then Return Nothing

            Dim newBitmap As New Bitmap(eAncho, eAlto)
            Dim g As Graphics = Graphics.FromImage(newBitmap)
            Dim colorMatrix As New ColorMatrix(New Single()() {New Single() {0.1F, 0.1F, 0.1F, 0, 0},
                                                               New Single() {0.3F, 0.3F, 0.3F, 0, 0},
                                                               New Single() {0.11F, 0.11F, 0.11F, 0, 0},
                                                               New Single() {0, 0, 0, 1, 0},
                                                               New Single() {0, 0, 0, 0, 1}})
            Dim attributes As New ImageAttributes()
            attributes.SetColorMatrix(colorMatrix)
            g.DrawImage(eImagen, New Rectangle(0, 0, eAncho, eAlto), 0, 0, eImagen.Width, eImagen.Height, GraphicsUnit.Pixel, attributes)
            g.Dispose()

            Return newBitmap
        End Function

        ''' <summary>
        ''' Obtiene una imagen con una transparencia
        ''' </summary>
        ''' <param name="eImagen">Imagen original</param>
        ''' <param name="eTransparencia">Nivel de transparencia a aplicar</param>
        ''' <param name="eAncho">Ancho de la imagen resultante</param>
        ''' <param name="eAlto">Alto de la imagen resultante</param>
        ''' <returns>Imagen con el nivel de transparencia y tamaños especificados</returns>
        Public Function obtenerTransparencia(ByVal eImagen As Bitmap, _
                                             ByVal eTransparencia As Single, _
                                             ByVal eAncho As Integer, _
                                             ByVal eAlto As Integer) As Bitmap
            If eImagen Is Nothing Then Return Nothing

            Dim newBitmap As New Bitmap(eAncho, eAlto)
            Dim g As Graphics = Graphics.FromImage(newBitmap)

            ' Initialize the color matrix.
            ' Note the value 0.8 in row 4, column 4.
            Dim matrixItems As Single()() = { _
               New Single() {1, 0, 0, 0, 0}, _
               New Single() {0, 1, 0, 0, 0}, _
               New Single() {0, 0, 1, 0, 0}, _
               New Single() {0, 0, 0, 0.8F, 0}, _
               New Single() {0, 0, 0, 0, 1}}

            Dim colorMatrix As New ColorMatrix(matrixItems)

            ' Create an ImageAttributes object and set its color matrix.
            Dim imageAtt As New ImageAttributes()
            imageAtt.SetColorMatrix( _
               colorMatrix, _
               ColorMatrixFlag.Default, _
               ColorAdjustType.Bitmap)

            ' Now draw the semitransparent bitmap image.
            Dim iWidth As Integer = eImagen.Width
            Dim iHeight As Integer = eImagen.Height

            ' Pass in the destination rectangle (2nd argument) and the x _
            ' coordinate (3rd argument), x coordinate (4th argument), width _
            ' (5th argument), and height (6th argument) of the source rectangle.
            g.DrawImage(eImagen, New Rectangle(0, 0, iWidth, iHeight), 0.0F, 0.0F, iWidth, iHeight, GraphicsUnit.Pixel, imageAtt)
            g.Dispose()

            Return newBitmap
        End Function

        ''' <summary>
        ''' Obtiene una imagen coloreada a partir de una matriz de color
        ''' </summary>
        ''' <param name="eImagen">Imagen original</param>
        ''' <param name="eMatriz">Matriz de color para colorear la imagen</param>
        ''' <returns>Imagen coloreada</returns>
        Public Function obtenerImagenColoreada(ByVal eImagen As Bitmap, _
                                               ByVal eMatriz() As Single) As Bitmap
            If eImagen Is Nothing Then Return Nothing

            Dim newBitmap As New Bitmap(eImagen.Width, eImagen.Height)
            Dim g As Graphics = Graphics.FromImage(newBitmap)

            Dim colorMatrix As New ColorMatrix(New Single()() {eMatriz,
                                                               eMatriz,
                                                               eMatriz,
                                                               New Single() {0, 0, 0, 1, 0},
                                                               New Single() {0, 0, 0, 0, 1}})

            Dim attributes As New ImageAttributes()
            attributes.SetColorMatrix(colorMatrix)
            g.DrawImage(eImagen, New Rectangle(0, 0, eImagen.Width, eImagen.Height), 0, 0, eImagen.Width, eImagen.Height, GraphicsUnit.Pixel, attributes)
            g.Dispose()
            Return newBitmap
        End Function
#End Region

#Region " STREAMS Y BASE 64 "
        ''' <summary>
        ''' Codifica la imagen que se le pasa como parámetro en su correspondiente cadena en Base64
        ''' </summary>
        ''' <param name="eImagen">Imagen a convertir a base64</param>
        ''' <param name="eFormato">Formato de la imagen, por defecto, png</param>
        ''' <returns>Cadena con la imagen convertida a base64</returns>
        Public Function imagen2Base64(ByVal eImagen As System.Drawing.Image, _
                                      Optional ByVal eFormato As System.Drawing.Imaging.ImageFormat = Nothing) As String
            Dim paraDevolver As String = ""

            If eFormato Is Nothing Then eFormato = System.Drawing.Imaging.ImageFormat.Png
            If eImagen Is Nothing Then Return ""

            ' Se realiza la conversión a base64
            Dim elMemoryStream As New IO.MemoryStream
            eImagen.Save(elMemoryStream, eFormato)
            paraDevolver = Convert.ToBase64String(elMemoryStream.ToArray)
            elMemoryStream.Close()

            Return paraDevolver
        End Function

        ''' <summary>
        ''' Decodifica una imagen codificada en base64 a una imagen
        ''' </summary>
        ''' <param name="eString">String con la imagen codificada en Base64</param>
        ''' <returns>La imagen para poder trabajar con ella</returns>
        Public Function base642Imagen(ByVal eString As String) As Image
            Try
                ' Se convierte a un array de bytes
                Dim imageBytes As Byte() = Convert.FromBase64String(eString)
                Dim ms As New MemoryStream(imageBytes, 0, imageBytes.Length)

                ' Y se convierte el array de bytes a la imagen
                ms.Write(imageBytes, 0, imageBytes.Length)
                Dim paraDevolver As Image = Image.FromStream(ms, True)
                Return paraDevolver
            Catch ex As Exception
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Convierte una imagen a un array de bytes
        ''' </summary>
        ''' <param name="eImagen">Imagen que se quiere transformar</param>
        ''' <param name="eFormato">Formato que se le va a aplicar a la imagen</param>
        ''' <returns>Array de bytes de la imagen</returns>
        Public Function imagen2Byte(ByVal eImagen As Image, _
                                    Optional ByVal eFormato As System.Drawing.Imaging.ImageFormat = Nothing) As Byte()
            If eImagen Is Nothing Then Return Nothing

            Dim imgStream As MemoryStream = New MemoryStream()
            If eFormato IsNot Nothing Then
                eImagen.Save(imgStream, eFormato)
            Else
                eImagen.Save(imgStream, System.Drawing.Imaging.ImageFormat.Png)
            End If
            imgStream.Close()

            Dim byteArray As Byte() = imgStream.ToArray()
            imgStream.Dispose()

            Return byteArray
        End Function

        ''' <summary>
        ''' Convierte un array de bytes en una imagen
        ''' </summary>
        ''' <param name="eByte">Array de bytes que contienen la imagen</param>
        ''' <returns>Imagen que se genera a partir de los bytes</returns>
        Public Function byte2Imagen(ByVal eByte As Byte()) As Image
            Dim ImageStream As MemoryStream
            Dim paraDevolver As Image
            Try
                If eByte.GetUpperBound(0) > 0 Then
                    ImageStream = New MemoryStream(eByte)
                    paraDevolver = Image.FromStream(ImageStream)
                Else
                    paraDevolver = Nothing
                End If
            Catch ex As System.Exception
                paraDevolver = Nothing
            End Try

            Return paraDevolver
        End Function
#End Region

#Region " GENERACION IMAGENES "

        ''' <summary>
        ''' Obtiene un cuadro del color y tamaño pasados por parámetros
        ''' </summary>            
        ''' <param name="eColor">Color para la creación del cuadro</param>
        ''' <param name="eAncho">Ancho del cuadro</param>
        ''' <param name="eAlto">Alto del cuadro</param>
        ''' <returns>Imagen generada a partir de los parámetros</returns>
        Public Function obtenerCuadroColor(ByVal eColor As Color, _
                                           ByVal eAncho As Integer, _
                                           ByVal eAlto As Integer) As Image
            Dim ParaDevolver As New Bitmap(eAncho, eAlto)
            Dim gr As Graphics = Graphics.FromImage(ParaDevolver)
            Dim Brocha As New SolidBrush(eColor)
            gr.FillRectangle(Brocha, 0, 0, eAncho, eAlto)
            Return ParaDevolver
        End Function

        ''' <summary>
        ''' Obtiene un cuiadro de texto a partir de los parámetros que se le pasan
        ''' </summary>
        ''' <param name="eAncho">Ancho del cuadro</param>
        ''' <param name="eAlto">Alto del cuadro</param>
        ''' <param name="eColorFondo">Color para el fondo</param>
        ''' <param name="eColorFuente">Cholor para la fuente</param>
        ''' <param name="eFuente">Fuente</param>
        ''' <param name="eTexto">Texto que se tiene que mostrar en el cuadro</param>
        ''' <returns>Imagen con las caracterísquetas que se le pasan</returns>
        Public Function obtenerCuadroTexto(ByVal eAncho As Integer, _
                                           ByVal eAlto As Integer, _
                                           ByVal eColorFondo As Color, _
                                           ByVal eColorFuente As Color, _
                                           ByVal eFuente As Font, _
                                           ByVal eTexto As String) As Image
            Dim ParaDevolver As New Bitmap(eAncho, eAlto)
            Dim gr As Graphics = Graphics.FromImage(ParaDevolver)

            If eColorFondo <> Color.Transparent Then
                Dim Brocha As New SolidBrush(eColorFondo)
                gr.FillRectangle(Brocha, 0, 0, eAncho, eAlto)
            Else
                gr.Clear(Color.Transparent)
            End If

            Dim brochaFuente As New SolidBrush(eColorFuente)
            gr.DrawString(eTexto, eFuente, brochaFuente, 0, 0)

            Return ParaDevolver
        End Function
#End Region

#Region " CAPTURAR VENTANAS / PANTALLA "
        ''' <summary>
        ''' Realiza una captura de pantalla de la ventana especificada
        ''' </summary>
        ''' <param name="eFormulario">Formulario del que se va a realizar la captura</param>
        ''' <returns>Imagen con la captura del formulario</returns>
        Public Function capturarImagenFormulario(ByVal eFormulario As Form) As Bitmap
            Dim Resultado As Bitmap = Nothing

            Using g As Graphics = Graphics.FromHwnd(eFormulario.Handle)
                Dim rc As Rectangle = eFormulario.DesktopBounds

                If g.VisibleClipBounds.Width > 0 AndAlso g.VisibleClipBounds.Height > 0 Then
                    Resultado = New Bitmap(rc.Width, rc.Height, g)

                    Using memoryGrahics As Graphics = Graphics.FromImage(Resultado)
                        memoryGrahics.CopyFromScreen(rc.X, rc.Y, 0, 0, rc.Size, CopyPixelOperation.SourceCopy)
                    End Using
                End If
            End Using

            Return Resultado
        End Function

        ''' <summary>
        ''' Realiza una captura de pantalla de la ventana especificada
        ''' </summary>
        ''' <returns>Captura de pantalla del escritorio</returns>
        Public Function CapturarImagenPantalla() As Bitmap
            Dim elResultado As Bitmap = Nothing

            Try
                Dim Limites As New Rectangle

                ' Obtengo los límites de todas las pantallas
                For Each Pantalla As Screen In Screen.AllScreens
                    If Pantalla.Bounds.X < Limites.X Then Limites.X = Pantalla.Bounds.X
                    If Pantalla.Bounds.Y < Limites.Y Then Limites.Y = Pantalla.Bounds.Y
                    If Pantalla.Bounds.Width + Pantalla.Bounds.X > Limites.Width Then Limites.Width = Pantalla.Bounds.Width + Pantalla.Bounds.X
                    If Pantalla.Bounds.Height + Pantalla.Bounds.Y > Limites.Height Then Limites.Height = Pantalla.Bounds.Height + Pantalla.Bounds.Y
                Next

                elResultado = New Bitmap(Limites.Width, Limites.Height)

                Using g As Graphics = Graphics.FromImage(elResultado)
                    g.CopyFromScreen(Point.Empty, Point.Empty, Limites.Size)
                End Using
            Catch ex As Exception
                If Log._LOG_ACTIVO Then Log.escribirLog("Se ha producido un error al realizar la captura de pantalla...", ex, New StackTrace(0, True))
                elResultado = Nothing
            End Try

            Return elResultado
        End Function

        ''' <summary>
        ''' Rezaliza una captura de la pantalla utilizando unas coordenadas y
        ''' dimensiones especificas
        ''' </summary>
        ''' <param name="ePosX">Posición X desde donde se va realizar la captura</param>
        ''' <param name="ePosY">Posición Y desde donde se va a realizar la captura</param>
        ''' <param name="eAncho">Ancho de la captura</param>
        ''' <param name="eAlto">Alto de la captura</param>
        ''' <returns>Bitmap con la captura de la pantalla</returns>
        Public Function CapturarImagenPantalla(ByVal ePosX As Long, _
                                               ByVal ePosY As Long, _
                                               ByVal eAncho As Long, _
                                               ByVal eAlto As Long) As Bitmap
            Return CapturarImagenPantalla(New Point(ePosX, ePosY), New Size(eAncho, eAlto))
        End Function

        ''' <summary>
        ''' Realiza una captura de pantalla utilizando una coordenada y 
        ''' una dimensión especificada
        ''' </summary>
        ''' <param name="eCoordenada">Coordenada donde se tiene que realizar la captura</param>
        ''' <param name="eTamanho">Tamaño de la captura</param>
        ''' <returns>Bitmap con la captura de pantalla</returns>
        Public Function CapturarImagenPantalla(ByVal eCoordenada As Point, _
                                               ByVal eTamanho As Size) As Bitmap
            Dim elResultado As Bitmap = Nothing
            Try
                elResultado = New Bitmap(eTamanho.Width, eTamanho.Height)
                elResultado.SetResolution(300, 300)

                Dim GR As Graphics = Graphics.FromImage(elResultado)
                GR.CopyFromScreen(eCoordenada.X, eCoordenada.Y, 0, 0, eTamanho)
            Catch ex As Exception
                If Log._LOG_ACTIVO Then Log.escribirLog("Se ha producido un error al realizar la captura de pantalla...", ex, New StackTrace(0, True))
                elResultado = Nothing
            End Try

            Return elResultado
        End Function
#End Region

    End Module
End Namespace