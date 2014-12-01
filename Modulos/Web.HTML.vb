﻿Imports System.Text.RegularExpressions
Imports System.Drawing

Namespace Web
    Namespace HTML
        Public Module modRecompilaWebHTML
            ''' <summary>
            ''' Realiza los cambios necesarios para convertir los caracteres especiales 
            ''' a formato HTML
            ''' </summary>
            ''' <param name="eHTML">HTML que se va a convertir</param>
            ''' <returns>HTML con todos los caracteres convertidos</returns>
            Public Function UTF2HTML(ByVal eHTML As String) As String
                ' Correcciones de acentos, eñes, etc etc
                eHTML = eHTML.Replace("á", "&aacute;")
                eHTML = eHTML.Replace("é", "&eacute;")
                eHTML = eHTML.Replace("í", "&iacute;")
                eHTML = eHTML.Replace("ó", "&oacute;")
                eHTML = eHTML.Replace("ú", "&uacute;")

                eHTML = eHTML.Replace("Á", "&Aacute;")
                eHTML = eHTML.Replace("É", "&Eacute;")
                eHTML = eHTML.Replace("Í", "&Iacute;")
                eHTML = eHTML.Replace("Ó", "&Oacute;")
                eHTML = eHTML.Replace("Ú", "&Uacute;")


                eHTML = eHTML.Replace("à", "&agrave;")
                eHTML = eHTML.Replace("è", "&egrave;")
                eHTML = eHTML.Replace("ì", "&igrave;")
                eHTML = eHTML.Replace("ò", "&ograve;")
                eHTML = eHTML.Replace("ù", "&ugrave;")

                eHTML = eHTML.Replace("À", "&Agrave;")
                eHTML = eHTML.Replace("È", "&Egrave;")
                eHTML = eHTML.Replace("Í", "&Igrave;")
                eHTML = eHTML.Replace("Ò", "&Ograve;")
                eHTML = eHTML.Replace("Ù", "&Ugrave;")


                eHTML = eHTML.Replace("â", "&acirc;")
                eHTML = eHTML.Replace("ê", "&ecirc;")
                eHTML = eHTML.Replace("î", "&icirc;")
                eHTML = eHTML.Replace("ô", "&ocirc;")
                eHTML = eHTML.Replace("û", "&ucirc;")

                eHTML = eHTML.Replace("Â", "&Acirc;")
                eHTML = eHTML.Replace("Ê", "&Ecirc;")
                eHTML = eHTML.Replace("Î", "&Icirc;")
                eHTML = eHTML.Replace("Ô", "&Ocirc;")
                eHTML = eHTML.Replace("Û", "&Ucirc;")


                eHTML = eHTML.Replace("ã", "&atilde;")
                eHTML = eHTML.Replace("Ã", "&Atilde;")
                eHTML = eHTML.Replace("õ", "&otilde;")
                eHTML = eHTML.Replace("Õ", "&Otilde;")
                eHTML = eHTML.Replace("ñ", "&ntilde;")
                eHTML = eHTML.Replace("Ñ", "&Ntilde;")


                eHTML = eHTML.Replace("ç", "&ccedil;")
                eHTML = eHTML.Replace("Ç", "&Ccedil;")


                eHTML = eHTML.Replace("ä", "&auml;")
                eHTML = eHTML.Replace("ë", "&euml;")
                eHTML = eHTML.Replace("ï", "&iuml;")
                eHTML = eHTML.Replace("ö", "&ouml;")
                eHTML = eHTML.Replace("ü", "&uuml;")

                eHTML = eHTML.Replace("Ä", "&Auml;")
                eHTML = eHTML.Replace("Ë", "&Euml;")
                eHTML = eHTML.Replace("Ï", "&Iuml;")
                eHTML = eHTML.Replace("Ö", "&Ouml;")
                eHTML = eHTML.Replace("Ü", "&Uuml;")


                eHTML = eHTML.Replace("«", "&#171;")
                eHTML = eHTML.Replace("»", "&#187;")

                eHTML = eHTML.Replace("º", "&deg;")
                eHTML = eHTML.Replace("ª", "&ordf;")

                Return eHTML
            End Function

            ''' <summary>
            ''' Convierte las expresiones HTML para caracteres especiales a los caracteres
            ''' que representan
            ''' </summary>
            ''' <param name="eHTML">Código HTML con los caracteres representados como expresiones</param>
            ''' <returns>Código HTML con los caracteres reales</returns>
            Public Function HTML2UTF(ByVal eHTML As String) As String
                ' Correcciones de acentos, eñes, etc etc
                eHTML = eHTML.Replace("&aacute;", "á")
                eHTML = eHTML.Replace("&eacute;", "é")
                eHTML = eHTML.Replace("&iacute;", "í")
                eHTML = eHTML.Replace("&oacute;", "ó")
                eHTML = eHTML.Replace("&uacute;", "ú")

                eHTML = eHTML.Replace("&Aacute;", "Á")
                eHTML = eHTML.Replace("&Eacute;", "É")
                eHTML = eHTML.Replace("&Iacute;", "Í")
                eHTML = eHTML.Replace("&Oacute;", "Ó")
                eHTML = eHTML.Replace("&Uacute;", "Ú")


                eHTML = eHTML.Replace("&agrave;", "à")
                eHTML = eHTML.Replace("&egrave;", "è")
                eHTML = eHTML.Replace("&igrave;", "ì")
                eHTML = eHTML.Replace("&ograve;", "ò")
                eHTML = eHTML.Replace("&ugrave;", "ù")

                eHTML = eHTML.Replace("&Agrave;", "À")
                eHTML = eHTML.Replace("&Egrave;", "È")
                eHTML = eHTML.Replace("&Igrave;", "Í")
                eHTML = eHTML.Replace("&Ograve;", "Ò")
                eHTML = eHTML.Replace("&Ugrave;", "Ù")


                eHTML = eHTML.Replace("&acirc;", "â")
                eHTML = eHTML.Replace("&ecirc;", "ê")
                eHTML = eHTML.Replace("&icirc;", "î")
                eHTML = eHTML.Replace("&ocirc;", "ô")
                eHTML = eHTML.Replace("&ucirc;", "û")

                eHTML = eHTML.Replace("&Acirc;", "Â")
                eHTML = eHTML.Replace("&Ecirc;", "Ê")
                eHTML = eHTML.Replace("&Icirc;", "Î")
                eHTML = eHTML.Replace("&Ocirc;", "Ô")
                eHTML = eHTML.Replace("&Ucirc;", "Û")


                eHTML = eHTML.Replace("&atilde;", "ã")
                eHTML = eHTML.Replace("&Atilde;", "Ã")
                eHTML = eHTML.Replace("&otilde;", "õ")
                eHTML = eHTML.Replace("&Otilde;", "Õ")
                eHTML = eHTML.Replace("&ntilde;", "ñ")
                eHTML = eHTML.Replace("&Ntilde;", "Ñ")


                eHTML = eHTML.Replace("&ccedil;", "ç")
                eHTML = eHTML.Replace("&Ccedil;", "Ç")


                eHTML = eHTML.Replace("&auml;", "ä")
                eHTML = eHTML.Replace("&euml;", "ë")
                eHTML = eHTML.Replace("&iuml;", "ï")
                eHTML = eHTML.Replace("&ouml;", "ö")
                eHTML = eHTML.Replace("&uuml;", "ü")

                eHTML = eHTML.Replace("&Auml;", "Ä")
                eHTML = eHTML.Replace("&Euml;", "Ë")
                eHTML = eHTML.Replace("&Iuml;", "Ï")
                eHTML = eHTML.Replace("&Ouml;", "Ö")
                eHTML = eHTML.Replace("&Uuml;", "Ü")


                eHTML = eHTML.Replace("&#171;", "«")
                eHTML = eHTML.Replace("&#187;", "»")

                eHTML = eHTML.Replace("&deg;", "º")
                eHTML = eHTML.Replace("&ordf;", "ª")

                Return eHTML
            End Function

            ''' <summary>
            ''' Convierte los caracteres UTF8 especiales a ANSI
            ''' </summary>
            ''' <param name="eTexto">Texto con los caracteres especiales UTF8</param>
            ''' <returns>Texto con los caracteres convertidos a ANSI</returns>
            Public Function UTF82ANSI(ByVal eTexto As String) As String
                eTexto = eTexto.Replace("¡", "Â¡")
                eTexto = eTexto.Replace("¢", "Â¢")
                eTexto = eTexto.Replace("£", "Â£")
                eTexto = eTexto.Replace("¤", "Â¤")
                eTexto = eTexto.Replace("¥", "Â¥")
                eTexto = eTexto.Replace("©", "Â©")
                eTexto = eTexto.Replace("ª", "Âª")
                eTexto = eTexto.Replace("«", "Â«")
                eTexto = eTexto.Replace("®", "Â®")
                eTexto = eTexto.Replace("°", "Â°")
                eTexto = eTexto.Replace("±", "Â±")
                eTexto = eTexto.Replace("²", "Â²")
                eTexto = eTexto.Replace("³", "Â³")
                eTexto = eTexto.Replace("´", "Â´")
                eTexto = eTexto.Replace("µ", "Âµ")
                eTexto = eTexto.Replace("·", "Â·")
                eTexto = eTexto.Replace("¹", "Â¹")
                eTexto = eTexto.Replace("º", "Âº")
                eTexto = eTexto.Replace("»", "Â»")
                eTexto = eTexto.Replace("¼", "Â¼")
                eTexto = eTexto.Replace("½", "Â½")
                eTexto = eTexto.Replace("¾", "Â¾")
                eTexto = eTexto.Replace("¿", "Â¿")
                eTexto = eTexto.Replace("À", "Ã€")
                eTexto = eTexto.Replace("Á", "Ã")
                eTexto = eTexto.Replace("Â", "Ã‚")
                eTexto = eTexto.Replace("Ã", "Ãƒ")
                eTexto = eTexto.Replace("Ä", "Ã„")
                eTexto = eTexto.Replace("Ç", "Ã‡")
                eTexto = eTexto.Replace("È", "Ãˆ")
                eTexto = eTexto.Replace("É", "Ã‰")
                eTexto = eTexto.Replace("Ê", "ÃŠ")
                eTexto = eTexto.Replace("Ë", "Ã‹")
                eTexto = eTexto.Replace("Ì", "ÃŒ")
                eTexto = eTexto.Replace("Í", "Ã")
                eTexto = eTexto.Replace("Î", "ÃŽ")
                eTexto = eTexto.Replace("Ï", "Ã")
                eTexto = eTexto.Replace("Ñ", "Ã‘")
                eTexto = eTexto.Replace("Ò", "Ã’")
                eTexto = eTexto.Replace("Ó", "Ã""")
                eTexto = eTexto.Replace("Õ", "Ã•")
                eTexto = eTexto.Replace("Ö", "Ã–")
                eTexto = eTexto.Replace("Ù", "Ã™")
                eTexto = eTexto.Replace("Ú", "Ãš")
                eTexto = eTexto.Replace("Û", "Ã›")
                eTexto = eTexto.Replace("Ü", "Ãœ")
                eTexto = eTexto.Replace("Ý", "Ã")
                eTexto = eTexto.Replace("à", "Ã ")
                eTexto = eTexto.Replace("á", "Ã¡")
                eTexto = eTexto.Replace("â", "Ã¢")
                eTexto = eTexto.Replace("ã", "Ã£")
                eTexto = eTexto.Replace("ä", "Ã¤")
                eTexto = eTexto.Replace("ç", "Ã§")
                eTexto = eTexto.Replace("è", "Ã¨")
                eTexto = eTexto.Replace("é", "Ã©")
                eTexto = eTexto.Replace("ê", "Ãª")
                eTexto = eTexto.Replace("ë", "Ã«")
                eTexto = eTexto.Replace("ì", "Ã")
                eTexto = eTexto.Replace("í", "Ã­")
                eTexto = eTexto.Replace("î", "Ã®")
                eTexto = eTexto.Replace("ï", "Ã¯")
                eTexto = eTexto.Replace("ñ", "Ã±")
                eTexto = eTexto.Replace("ò", "Ã²")
                eTexto = eTexto.Replace("ó", "Ã³")
                eTexto = eTexto.Replace("ô", "Ã´")
                eTexto = eTexto.Replace("õ", "Ãµ")
                eTexto = eTexto.Replace("ö", "Ã")
                eTexto = eTexto.Replace("ù", "Ã¹")
                eTexto = eTexto.Replace("ú", "Ãº")
                eTexto = eTexto.Replace("û", "Ã»")
                eTexto = eTexto.Replace("ü", "Ã¼")
                eTexto = eTexto.Replace("ÿ", "Ã¿")
                Return eTexto
            End Function

            ''' <summary>
            ''' Convierte los caracteres especiales ANSI a caracteres UTF8
            ''' </summary>
            ''' <param name="eTexto">Texto con caracteres especiales ANSI</param>
            ''' <returns>Texto con los caracteres especiales convertidos a UTF8</returns>
            Public Function ANSI2UTF8(ByVal eTexto As String) As String
                eTexto = eTexto.Replace("Â¡", "¡")
                eTexto = eTexto.Replace("Â¢", "¢")
                eTexto = eTexto.Replace("Â£", "£")
                eTexto = eTexto.Replace("Â¤", "¤")
                eTexto = eTexto.Replace("Â¥", "¥")
                eTexto = eTexto.Replace("Â©", "©")
                eTexto = eTexto.Replace("Âª", "ª")
                eTexto = eTexto.Replace("Â«", "«")
                eTexto = eTexto.Replace("Â®", "®")
                eTexto = eTexto.Replace("Â°", "°")
                eTexto = eTexto.Replace("Â±", "±")
                eTexto = eTexto.Replace("Â²", "²")
                eTexto = eTexto.Replace("Â³", "³")
                eTexto = eTexto.Replace("Â´", "´")
                eTexto = eTexto.Replace("Âµ", "µ")
                eTexto = eTexto.Replace("Â·", "·")
                eTexto = eTexto.Replace("Â¹", "¹")
                eTexto = eTexto.Replace("Âº", "º")
                eTexto = eTexto.Replace("Â»", "»")
                eTexto = eTexto.Replace("Â¼", "¼")
                eTexto = eTexto.Replace("Â½", "½")
                eTexto = eTexto.Replace("Â¾", "¾")
                eTexto = eTexto.Replace("Â¿", "¿")
                eTexto = eTexto.Replace("Ã€", "À")
                eTexto = eTexto.Replace("Ã", "Á")
                eTexto = eTexto.Replace("Ã‚", "Â")
                eTexto = eTexto.Replace("Ãƒ", "Ã")
                eTexto = eTexto.Replace("Ã„", "Ä")
                eTexto = eTexto.Replace("Ã‡", "Ç")
                eTexto = eTexto.Replace("Ãˆ", "È")
                eTexto = eTexto.Replace("Ã‰", "É")
                eTexto = eTexto.Replace("ÃŠ", "Ê")
                eTexto = eTexto.Replace("Ã‹", "Ë")
                eTexto = eTexto.Replace("ÃŒ", "Ì")
                eTexto = eTexto.Replace("Ã", "Í")
                eTexto = eTexto.Replace("ÃŽ", "Î")
                eTexto = eTexto.Replace("Ã", "Ï")
                eTexto = eTexto.Replace("Ã‘", "Ñ")
                eTexto = eTexto.Replace("Ã’", "Ò")
                eTexto = eTexto.Replace("Ã""", "Ó")
                eTexto = eTexto.Replace("Ã•", "Õ")
                eTexto = eTexto.Replace("Ã–", "Ö")
                eTexto = eTexto.Replace("Ã™", "Ù")
                eTexto = eTexto.Replace("Ãš", "Ú")
                eTexto = eTexto.Replace("Ã›", "Û")
                eTexto = eTexto.Replace("Ãœ", "Ü")
                eTexto = eTexto.Replace("Ã", "Ý")
                eTexto = eTexto.Replace("Ã ", "à")
                eTexto = eTexto.Replace("Ã¡", "á")
                eTexto = eTexto.Replace("Ã¢", "â")
                eTexto = eTexto.Replace("Ã£", "ã")
                eTexto = eTexto.Replace("Ã¤", "ä")
                eTexto = eTexto.Replace("Ã§", "ç")
                eTexto = eTexto.Replace("Ã¨", "è")
                eTexto = eTexto.Replace("Ã©", "é")
                eTexto = eTexto.Replace("Ãª", "ê")
                eTexto = eTexto.Replace("Ã«", "ë")
                eTexto = eTexto.Replace("Ã", "ì")
                eTexto = eTexto.Replace("Ã­", "í")
                eTexto = eTexto.Replace("Ã®", "î")
                eTexto = eTexto.Replace("Ã¯", "ï")
                eTexto = eTexto.Replace("Ã±", "ñ")
                eTexto = eTexto.Replace("Ã²", "ò")
                eTexto = eTexto.Replace("Ã³", "ó")
                eTexto = eTexto.Replace("Ã´", "ô")
                eTexto = eTexto.Replace("Ãµ", "õ")
                eTexto = eTexto.Replace("Ã", "ö")
                eTexto = eTexto.Replace("Ã¹", "ù")
                eTexto = eTexto.Replace("Ãº", "ú")
                eTexto = eTexto.Replace("Ã»", "û")
                eTexto = eTexto.Replace("Ã¼", "ü")
                eTexto = eTexto.Replace("Ã¿", "ÿ")
                Return eTexto
            End Function

            ''' <summary>
            ''' A partir de una página HTML que tiene imágenes no incrustadas,
            ''' genera el mismo HTML cambiando las imágenes absolutas por imágenes PNG incrustadas
            ''' en base64
            ''' </summary>
            ''' <param name="eHTML">Código HTML original</param>
            ''' <param name="eRutaDirectorioImagenes">Ruta base donde se encuentran las imágenes</param>
            ''' <returns>Código HTML con las imágenes incrustadas</returns>
            Public Function HTML2HTML_IMGBase64(ByVal eHTML As String, _
                                                ByVal eRutaDirectorioImagenes As String) As String
                ' Si la ruta de las imágenes no termina con \ esta se añade
                If Not String.IsNullOrEmpty(eRutaDirectorioImagenes) AndAlso Not eRutaDirectorioImagenes.EndsWith("\") Then eRutaDirectorioImagenes &= "\"

                ' Se crea la plantilla modificada, que será identica a la original cambiando los ficheros por 
                ' las imágenes en BASE64
                Dim PlantillaModificada As String = eHTML

                ' Se aplica una expresión regular para obtener los nombres de los ficheros que se incluyen
                ' en la firma y después pasalos a BASE64
                Dim elPatron As String = "(?i:<img.*?src=""(.*?)"".*?>)"
                Dim losResultados As MatchCollection = Regex.Matches(eHTML, elPatron)
                For Each unResultado As Match In losResultados
                    ' Se obtiene la ruta del fichero
                    Dim nombreFichero As String = Regex.Match(unResultado.Groups(0).Value, "<img.+?src=[""'](.+?)[""'].*?>", RegexOptions.IgnoreCase).Groups(1).Value
                    Dim rutaFichero As String = eRutaDirectorioImagenes & nombreFichero

                    ' Se comprueba que exista el fichero en la ruta, si existe se carga en una imagen
                    If IO.File.Exists(rutaFichero) Then
                        Dim laImagen As Image = Image.FromFile(rutaFichero)
                        Dim laImagenBase64 As String = "data:image/png;base64," & Imagenes.imagen2Base64(laImagen, System.Drawing.Imaging.ImageFormat.Png)

                        ' Se reemplaza la imagen original por la imagen en base 64
                        Dim lasCadenasBusqueda As New List(Of String)
                        With lasCadenasBusqueda
                            .Add("src=""" & nombreFichero & """")
                            .Add("src =""" & nombreFichero & """")
                            .Add("src= """ & nombreFichero & """")
                            .Add("src = """ & nombreFichero & """")
                            .Add("src='" & nombreFichero & "'")
                            .Add("src ='" & nombreFichero & "'")
                            .Add("src= '" & nombreFichero & "'")
                            .Add("src = '" & nombreFichero & "'")
                        End With
                        Dim cadenaReemplazo As String = "src='" & laImagenBase64 & "'"
                        For Each unaCadenaBusqueda As String In lasCadenasBusqueda
                            PlantillaModificada = PlantillaModificada.Replace(unaCadenaBusqueda, cadenaReemplazo)
                        Next
                    End If
                Next

                Return PlantillaModificada
            End Function
        End Module
    End Namespace
End Namespace