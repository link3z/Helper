Imports System.Text.RegularExpressions
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
            ''' Convierte los caracteres ISO a ASCII
            ''' </summary>
            ''' <param name="eHTML">HTML al que se le van a cambiar los caracteres ISO a ASCII</param>
            ''' <returns>HTML con la conversión realizada</returns>
            Public Function ISO2ASCII(ByVal eHTML As String) As String

                Dim strFrom As String() = New String(130) {}
                Dim strTo As String() = New String(130) {}
                strFrom(1) = "À"
                strTo(1) = "À"
                'capital A, grave accent       
                strFrom(2) = "à"
                strTo(2) = "à"
                'small a, grave accent        
                strFrom(3) = "Á"
                strTo(3) = "Á"
                'capital A, acute accent      
                strFrom(4) = "á"
                strTo(4) = "á"
                'small a, acute accent        
                strFrom(5) = "Â"
                strTo(5) = "Â"
                'capital A, circumflex        
                strFrom(6) = "â"
                strTo(6) = "â"
                'small a, circumflex          
                strFrom(7) = "Ã"
                strTo(7) = "Ã"
                'capital A, tilde             
                strFrom(8) = "ã"
                strTo(8) = "ã"
                'small a, tilde               
                strFrom(9) = "Ä"
                strTo(9) = "Ä"
                'capital A, diæresis/umlaut   
                strFrom(10) = "ä"
                strTo(10) = "ä"
                'small a, diæresis/umlaut     
                strFrom(11) = "Å"
                strTo(11) = "Å"
                'capital A, ring              
                strFrom(12) = "å"
                strTo(12) = "å"
                'small a, ring                
                strFrom(13) = "Æ"
                strTo(13) = "Æ"
                'capital AE ligature          
                strFrom(14) = "æ"
                strTo(14) = "æ"
                'small ae ligature            
                strFrom(15) = "Ç"
                strTo(15) = "Ç"
                'capital C, cedilla           
                strFrom(16) = "ç"
                strTo(16) = "ç"
                'small c, cedilla             
                strFrom(17) = "È"
                strTo(17) = "È"
                'capital E, grave accent      
                strFrom(18) = "è"
                strTo(18) = "è"
                'small e, grave accent        
                strFrom(19) = "É"
                strTo(19) = "É"
                'capital E, acute accent      
                strFrom(20) = "é"
                strTo(20) = "é"
                'small e, acute accent        
                strFrom(21) = "Ê"
                strTo(21) = "Ê"
                'capital E, circumflex        
                strFrom(22) = "ê"
                strTo(22) = "ê"
                'small e, circumflex          
                strFrom(23) = "Ë"
                strTo(23) = "Ë"
                'capital E, diæresis/umlaut   
                strFrom(24) = "ë"
                strTo(24) = "ë"
                'small e, diæresis/umlaut     
                strFrom(25) = "Ì"
                strTo(25) = "Ì"
                'capital I, grave accent      
                strFrom(26) = "ì"
                strTo(26) = "ì"
                'small i, grave accent        
                strFrom(27) = "Í"
                strTo(27) = "Í"
                'capital I, acute accent      
                strFrom(28) = "í"
                strTo(28) = "í"
                'small i, acute accent        
                strFrom(29) = "Î"
                strTo(29) = "Î"
                'capital I, circumflex        
                strFrom(30) = "î"
                strTo(30) = "î"
                'small i, circumflex          
                strFrom(31) = "Ï"
                strTo(31) = "Ï"
                'capital I, diæresis/umlaut   
                strFrom(32) = "ï"
                strTo(32) = "ï"
                'small i, diæresis/umlaut  
                strFrom(33) = "Ð"
                strTo(33) = "Ð"
                'capital Eth, Icelandic
                strFrom(34) = "ð"
                strTo(34) = "ð"
                'small eth, Icelandic
                strFrom(35) = "Ñ"
                strTo(35) = "Ñ"
                'capital N, tilde        
                strFrom(36) = "ñ"
                strTo(36) = "ñ"
                'small n, tilde               
                strFrom(37) = "Ò"
                strTo(37) = "Ò"
                'capital O, grave accent      
                strFrom(38) = "ò"
                strTo(38) = "ò"
                'small o, grave accent             
                strFrom(39) = "Ó"
                strTo(39) = "Ó"
                'capital O, acute accent      
                strFrom(40) = "ó"
                strTo(40) = "ó"
                'small o, acute accent        
                strFrom(41) = "Ô"
                strTo(41) = "Ô"
                'capital O, circumflex   
                strFrom(42) = "ô"
                strTo(42) = "ô"
                'small o, circumflex            
                strFrom(43) = "Õ"
                strTo(43) = "Õ"
                'capital O, tilde             
                strFrom(44) = "õ"
                strTo(44) = "õ"
                'small o, tilde               
                strFrom(45) = "Ö"
                strTo(45) = "Ö"
                'capital O, diæresis/umlaut 
                strFrom(46) = "ö"
                strTo(46) = "ö"
                'small o, diæresis/umlaut   
                strFrom(47) = "Ø"
                strTo(47) = "Ø"
                'capital O, slash                   
                strFrom(48) = "ø"
                strTo(48) = "ø"
                'small o, slash          
                strFrom(49) = "Ù"
                strTo(49) = "Ù"
                'capital U, grave accent           
                strFrom(50) = "ù"
                strTo(50) = "ù"
                'small u, grave accent        
                strFrom(51) = "Ú"
                strTo(51) = "Ú"
                'capital U, acute accent      
                strFrom(52) = "ú"
                strTo(52) = "ú"
                'small u, acute accent        
                strFrom(53) = "Û"
                strTo(53) = "Û"
                'capital U, circumflex          
                strFrom(54) = "û"
                strTo(54) = "û"
                'small u, circumflex            
                strFrom(55) = "Ü"
                strTo(55) = "Ü"
                'capital U, diæresis/umlaut 
                strFrom(56) = "ü"
                strTo(56) = "ü"
                'small u, diæresis/umlaut      
                strFrom(57) = "Ý"
                strTo(57) = "Ý"
                'capital Y, acute accent      
                strFrom(58) = "ý"
                strTo(58) = "ý"
                'small y, acute accent        
                strFrom(59) = "Þ"
                strTo(59) = "Þ"
                'capital Thorn, Icelandic       
                strFrom(60) = "þ"
                strTo(60) = "þ"
                'small thorn, Icelandic         
                strFrom(61) = "ß"
                strTo(61) = "ß"
                'small sharp s, German sz           
                strFrom(62) = "ÿ"
                strTo(62) = "ÿ"
                'small y, diæresis/umlaut 
                strFrom(63) = " "
                strTo(63) = " "
                'non-breaking space          
                strFrom(64) = "¡"
                strTo(64) = "¡"
                'inverted exclamation mark   
                strFrom(65) = "¢"
                strTo(65) = "¢"
                'cent sign                   
                strFrom(66) = "£"
                strTo(66) = "£"
                'pound sign                  
                strFrom(67) = "¤"
                strTo(67) = "¤"
                'general currency sign       
                strFrom(68) = "¥"
                strTo(68) = "¥"
                'yen sign                    
                strFrom(69) = "¦"
                strTo(69) = "¦"
                'broken [vertical] bar       
                strFrom(70) = "§"
                strTo(70) = "§"
                'section sign                
                strFrom(71) = "¨"
                strTo(71) = "¨"
                'umlaut/dieresis             
                strFrom(72) = "©"
                strTo(72) = "©"
                'copyright sign              
                strFrom(73) = "ª"
                strTo(73) = "ª"
                'ordinal indicator, fem      
                strFrom(74) = "«"
                strTo(74) = "«"
                'angle quotation mark, left  
                strFrom(75) = "¬"
                strTo(75) = "¬"
                'not sign                    
                strFrom(76) = "­"
                strTo(76) = "­"
                'soft hyphen                 
                strFrom(77) = "®"
                strTo(77) = "®"
                'registered sign             
                strFrom(78) = "¯"
                strTo(78) = "¯"
                'macron                      
                strFrom(79) = "°"
                strTo(79) = "°"
                'degree sign                 
                strFrom(80) = " "
                strTo(80) = " "
                'non-breaking space          
                strFrom(81) = "¡"
                strTo(81) = "¡"
                'inverted exclamation mark   
                strFrom(82) = "¢"
                strTo(82) = "¢"
                'cent sign                   
                strFrom(83) = "£"
                strTo(83) = "£"
                'pound sign                  
                strFrom(84) = "¤"
                strTo(84) = "¤"
                'general currency sign       
                strFrom(85) = "¥"
                strTo(85) = "¥"
                'yen sign                    
                strFrom(86) = "¦"
                strTo(86) = "¦"
                'broken [vertical] bar       
                strFrom(87) = "§"
                strTo(87) = "§"
                'section sign                
                strFrom(88) = "¨"
                strTo(88) = "¨"
                'umlaut/dieresis             
                strFrom(89) = "©"
                strTo(89) = "©"
                'copyright sign              
                strFrom(90) = "ª"
                strTo(90) = "ª"
                'ordinal indicator, fem      
                strFrom(91) = "«"
                strTo(91) = "«"
                'angle quotation mark, left  
                strFrom(92) = "¬"
                strTo(92) = "¬"
                'not sign                    
                strFrom(93) = "­"
                strTo(93) = "­"
                'soft hyphen                 
                strFrom(94) = "®"
                strTo(94) = "®"
                'registered sign             
                strFrom(95) = "¯"
                strTo(95) = "¯"
                'macron                      
                strFrom(96) = "°"
                strTo(96) = "°"
                'degree sign                 
                strFrom(97) = "±"
                strTo(97) = "±"
                'plus-or-minus sign          
                strFrom(98) = "²"
                strTo(98) = "²"
                'superscript two          
                strFrom(99) = "³"
                strTo(99) = "³"
                'superscript three        
                strFrom(100) = "´"
                strTo(100) = "´"
                'acute accent             
                strFrom(101) = "µ"
                strTo(101) = "µ"
                'micro sign                
                strFrom(102) = "¶"
                strTo(102) = "¶"
                'pilcrow [paragraph sign] 
                strFrom(103) = "·"
                strTo(103) = "·"
                'middle dot               
                strFrom(104) = "¸"
                strTo(104) = "¸"
                'cedilla                  
                strFrom(105) = "¹"
                strTo(105) = "¹"
                'superscript one          
                strFrom(106) = "º"
                strTo(106) = "º"
                'ordinal indicator, male  
                strFrom(107) = "»"
                strTo(107) = "»"
                'angle quotation mark, right   
                strFrom(108) = "¼"
                strTo(108) = "¼"
                'fraction one-quarter          
                strFrom(109) = "½"
                strTo(109) = "½"
                'fraction one-half             
                strFrom(110) = "¾"
                strTo(110) = "¾"
                'fraction three-quarters       
                strFrom(111) = "¿"
                strTo(111) = "¿"
                'inverted question mark        
                strFrom(112) = "×"
                strTo(112) = "×"
                'multiply sign                 
                strFrom(113) = "&div;"
                strTo(113) = "÷"
                'division sign             
                strFrom(114) = "±"
                strTo(114) = "±"
                'plus-or-minus sign          
                strFrom(115) = "²"
                strTo(115) = "²"
                'superscript two          
                strFrom(116) = "³"
                strTo(116) = "³"
                'superscript three        
                strFrom(117) = "´"
                strTo(117) = "´"
                'acute accent             
                strFrom(118) = "µ"
                strTo(118) = "µ"
                'micro sign                
                strFrom(119) = "¶"
                strTo(119) = "¶"
                'pilcrow [paragraph sign] 
                strFrom(120) = "·"
                strTo(120) = "·"
                'middle dot               
                strFrom(121) = "¸"
                strTo(121) = "¸"
                'cedilla                  
                strFrom(122) = "¹"
                strTo(122) = "¹"
                'superscript one          
                strFrom(123) = "º"
                strTo(123) = "º"
                'ordinal indicator, male  
                strFrom(124) = "»"
                strTo(124) = "»"
                'angle quotation mark, right   
                strFrom(125) = "¼"
                strTo(125) = "¼"
                'fraction one-quarter          
                strFrom(126) = "½"
                strTo(126) = "½"
                'fraction one-half             
                strFrom(127) = "¾"
                strTo(127) = "¾"
                'fraction three-quarters       
                strFrom(128) = "¿"
                strTo(128) = "¿"
                'inverted question mark        
                strFrom(129) = "×"
                strTo(129) = "×"
                'multiply sign                 
                strFrom(130) = "÷"
                strTo(130) = "÷"
                'division sign             
                Dim i As Integer
                For i = 1 To strFrom.Length - 1
                    eHTML = Regex.Replace(eHTML, strFrom(i), strTo(i), RegexOptions.IgnoreCase Or RegexOptions.Singleline)
                Next
                Return (eHTML)
            End Function

            ''' <summary>
            ''' A partir de una página HTML que tiene imágenes no incrustadas,
            ''' genera el mismo HTML cambiando las imágenes absolutas por imágenes PNG incrustadas
            ''' en base64
            ''' </summary>
            ''' <param name="eHTML">Código HTML original</param>
            ''' <returns>Código HTML con las imágenes incrustadas</returns>
            Public Function HTML2HTML_IMG64(ByVal eHTML As String) As String
                ' Se crea la plantilla modificada, que será identica a la original cambiando los ficheros por 
                ' las imágenes en BASE64
                Dim PlantillaModificada As String = eHTML

                ' Se aplica una expresión regular para obtener los nombres de los ficheros que se incluyen
                ' en la firma y después pasalos a BASE64
                Dim elPatron As String = "(?i:<img.*?src=""(.*?)"".*?>)"
                Dim losResultados As MatchCollection = Regex.Matches(eHTML, elPatron)
                For Each unResultado As Match In losResultados
                    ' Se obtiene la url donde está la imagen
                    Dim urlFichero As String = Regex.Match(unResultado.Groups(0).Value, "<img.+?src=[""'](.+?)[""'].*?>", RegexOptions.IgnoreCase).Groups(1).Value

                    If Not urlFichero.ToLower.StartsWith("data:") Then
                        ' Se trata de obtener la imagen desde la url
                        Dim laImagen As Image = Imagenes.obtenerImagenHTTP2Image(urlFichero)

                        ' Si se pudo obtener la imagen esta se convierta a base64
                        If laImagen IsNot Nothing Then
                            Dim laImagenBase64 As String = "data:image/png;base64," & Imagenes.imagen2Base64(laImagen, System.Drawing.Imaging.ImageFormat.Png)

                            ' Se reemplaza la imagen original por la imagen en base 64
                            Dim lasCadenasBusqueda As New List(Of String)
                            With lasCadenasBusqueda
                                .Add("src=""" & urlFichero & """")
                                .Add("src =""" & urlFichero & """")
                                .Add("src= """ & urlFichero & """")
                                .Add("src = """ & urlFichero & """")
                                .Add("src='" & urlFichero & "'")
                                .Add("src ='" & urlFichero & "'")
                                .Add("src= '" & urlFichero & "'")
                                .Add("src = '" & urlFichero & "'")
                            End With
                            Dim cadenaReemplazo As String = "src='" & laImagenBase64 & "'"
                            For Each unaCadenaBusqueda As String In lasCadenasBusqueda
                                PlantillaModificada = PlantillaModificada.Replace(unaCadenaBusqueda, cadenaReemplazo)
                            Next
                        End If
                    End If
                Next

                Return PlantillaModificada
            End Function

            ''' <summary>
            ''' Obtine una imagen desde una URL especificada
            ''' </summary>
            ''' <param name="eURL">Dirección URL de la imagen</param>
            ''' <returns>Imagen obtenida de la dirección URL</returns>
            Public Function obtenerImagenHTTP2Image(ByVal eURL As String) As Image
                Return Imagenes.obtenerImagenHTTP2Image(eURL)
            End Function

            ''' <summary>
            ''' Obtine una imagen desde una URL especificada devolviendo el array de bytes que la representa
            ''' </summary>
            ''' <param name="eURL">Dirección URL de la imagen</param>
            ''' <returns>Imagen obtenida de la dirección URL</returns>
            Public Function obtenerImagenHTTP2Byte(ByVal eURL As String) As Byte()
                Return Imagenes.obtenerImagenHTTP2Byte(eURL)
            End Function
        End Module
    End Namespace
End Namespace
