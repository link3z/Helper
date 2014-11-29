Imports System.Text.RegularExpressions

' La mayor parte de estos algoritmos se encuentran en: http://mvp-access.es/softjaen/vbnet/funciones/dc/index.htm
' Por si hay que incluir nuevos, comprobar antes si estos ya se encuentran codificados

Namespace Validadores
    Public Module modRecomilaValidadores
#Region " DECLARACIONES "
        Public Const _TEXTO_OBLIGATORIO As String = "El campo es obligatorio y no puede quedar en blanco."
        Public Const _EMAIL_NO_VALIDO As String = "El correo electrónico introducido no parece ser válido."
        Public Const _COMBO_OBLIGATORIO As String = "Es necesario que escoja una opción."
        Public Const _CHECKBOX_OBLIGATORIO As String = "Es necesario marcar esta opción."
#End Region

#Region " ENTRADA DE DATOS"
        ''' <summary>
        ''' Comprueba que el email que se le pasa por parametro sea correcto.
        ''' Se pueden validar más de un correo electrónico. El caracter de separación
        ''' entre estos va a ser el ;
        ''' </summary>
        ''' <param name="eEmail">Email o lista de emails separados por ; a ser comprobados</param>
        ''' <returns>True o False dependiendo de si es correcto o no</returns>
        Public Function validarEmail(ByVal eEmail As String) As Boolean
            If eEmail Is Nothing Then Return False
            Dim Patron As String = "^[a-zA-Z0-9][\w\.-]*[a-zA-Z0-9]@[a-zA-Z0-9][\w\.-]*[a-zA-Z0-9]\.[a-zA-Z][a-zA-Z\.]*[a-zA-Z]$"

            Dim ListaCorreos() As String = eEmail.Split(";")
            Dim ComprobacionCorrecta As Boolean = True
            If ListaCorreos.Length > 0 Then
                For Each UnEmail As String In ListaCorreos
                    If UnEmail.Trim <> "" Then
                        Dim ValidadorEmail As Match = Regex.Match(UnEmail, Patron)
                        ComprobacionCorrecta = ValidadorEmail.Success
                        If Not ComprobacionCorrecta Then Exit For
                    End If
                Next
            Else
                Return False
            End If

            Return ComprobacionCorrecta
        End Function

        ''' <summary>
        ''' Comprueba si la URL que se le pasa por parámetro está bien formada
        ''' </summary>
        ''' <param name="eURL">URL que se quiere validar</param>
        ''' <returns>True o False dependiendo de si es válida</returns>
        Public Function validarURL(ByVal eURL As String) As Boolean
            If eURL Is Nothing Then Return False
            Dim Patron As String = "((https?|ftp|gopher|telnet|file|notes|ms-help):((//)|(\\\\))+[\w\d:#@%/;$()~_?\+-=\\\.&]*)"
            Dim ValidadorURL As Match = Regex.Match(eURL, Patron)
            Return (ValidadorURL.Success)
        End Function

        ''' <summary>
        ''' Captura el evento keyPress y sólo permite que se ejecunten caracteres numéricos o de control
        ''' </summary>
        ''' <param name="sender">Control que desencadena el evento</param>
        ''' <param name="e">Parámetros pasados a la función</param>
        Public Sub SoloNumeros(ByVal sender As System.Object, ByRef e As System.Windows.Forms.KeyPressEventArgs)
            SoloNumeros(sender, e, True)
        End Sub

        ''' <summary>
        ''' Captura el evento keyPress y sólo permite que se ejecunten caracteres numéricos o de control
        ''' </summary>
        ''' <param name="sender">Control que desencadena el evento</param>
        ''' <param name="e">Parámetros pasados a la función</param>
        ''' <param name="PermitirNegativos">Permite la introducción de números negativos</param>
        Public Sub soloNumeros(ByVal sender As System.Object, ByRef e As System.Windows.Forms.KeyPressEventArgs, ByVal PermitirNegativos As Boolean)
            Try
                If e.KeyChar = "-" AndAlso Not sender.Text.StartsWith("-") AndAlso PermitirNegativos Then
                    e.Handled = True
                    sender.Text = "-" & sender.Text

                    ' Se situa el cursor al final del text
                    sender.SelectionStart = sender.Text.Length
                    sender.ScrollToCaret()
                ElseIf e.KeyChar = "+" AndAlso sender.Text.StartsWith("-") AndAlso PermitirNegativos Then
                    e.Handled = True
                    sender.Text = sender.Text.Replace("-", "")

                    ' Se situa el cursor al final del text
                    sender.SelectionStart = sender.Text.Length
                    sender.ScrollToCaret()
                ElseIf Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsNumber(e.KeyChar) Then
                    e.Handled = True
                End If
            Catch ex As System.Exception
                e.Handled = False
            End Try
        End Sub

        ''' <summary>
        ''' Permite introducir números con separador de decimales, verificando que sólo exista un único separador
        ''' </summary>
        ''' <param name="sender">Control que desencadena el evento</param>
        ''' <param name="e">Parámetros pasados a la función</param>
        Public Sub soloNumerosYDecimales(ByVal sender As Object, ByRef e As System.Windows.Forms.KeyPressEventArgs)
            Try
                If e.KeyChar = "." Then e.KeyChar = ","

                If InStr(sender.Text, ",") > 0 And e.KeyChar = "," Then
                    e.Handled = True
                ElseIf e.KeyChar = "," AndAlso String.IsNullOrEmpty(sender.Text) Then
                    e.Handled = True
                    sender.Text = "0,"

                    ' Se situa el cursor al final del text
                    sender.SelectionStart = 2
                    sender.ScrollToCaret()
                ElseIf e.KeyChar <> "," Then
                    SoloNumeros(sender, e)
                End If
            Catch ex As System.Exception
                e.Handled = False
            End Try
        End Sub

        ''' <summary>
        ''' Permite introducir números con separador de decimales, verificando que sólo exista un único separador
        ''' </summary>
        ''' <param name="sender">Control que desencadena el evento</param>
        ''' <param name="e">Parámetros pasados a la función</param>
        ''' <param name="PermitirNegativos">Permite la introducción de números negativos</param>
        Public Sub soloNumerosYDecimales(ByVal sender As Object, ByRef e As System.Windows.Forms.KeyPressEventArgs, _
                                         ByVal PermitirNegativos As Boolean)
            Try
                If e.KeyChar = "." Then e.KeyChar = ","

                If InStr(sender.Text, ",") > 0 And e.KeyChar = "," Then
                    e.Handled = True
                ElseIf e.KeyChar = "," AndAlso String.IsNullOrEmpty(sender.Text) Then
                    e.Handled = True
                    sender.Text = "0,"

                    ' Se situa el cursor al final del text
                    sender.SelectionStart = 2
                    sender.ScrollToCaret()
                ElseIf e.KeyChar <> "," Then
                    SoloNumeros(sender, e)
                End If
            Catch ex As System.Exception
                e.Handled = False
            End Try
        End Sub

        ''' <summary>
        ''' Permite introducir solamente letras en un control que tenga evento KeyPress
        ''' </summary>
        ''' <param name="sender">Objeto que envía el evento</param>
        ''' <param name="e">Parámetros del evento</param>
        ''' <param name="eSoloMayusculas">Indica si solamente se permiten mayusculas</param>
        ''' <param name="eLetrasEspeciales">Indica si se permiten letras especiales (ñ y ç)</param>
        Public Sub soloLetras(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs, _
                              ByVal eSoloMayusculas As Boolean,
                              ByVal eLetrasEspeciales As Boolean)
            If eSoloMayusculas Then
                If eLetrasEspeciales Then
                    If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZÑÇ" & Chr(8), e.KeyChar) = 0 Then e.KeyChar = ""
                Else
                    If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), e.KeyChar) = 0 Then e.KeyChar = ""
                End If
            Else
                If eLetrasEspeciales Then
                    If InStr(1, "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZçÇñÑ" & Chr(8), e.KeyChar) = 0 Then e.KeyChar = ""
                Else
                    If InStr(1, "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), e.KeyChar) = 0 Then e.KeyChar = ""
                End If
            End If
        End Sub

        ''' <summary>
        ''' Solamente permite la entrada de letras y números en un TextBox mediante el KeyPress
        ''' </summary>
        ''' <param name="sender">Objeto que envía el evento</param>
        ''' <param name="e">Parámetros del evento</param>
        ''' <param name="eSoloMayusculas">Indica si solamente se permiten mayusculas</param>
        ''' <param name="eLetrasEspeciales">Indica si se permiten letras especiales (ñ y ç)</param>
        Public Sub SoloLetrasYNumeros(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs, _
                                      ByVal eSoloMayusculas As Boolean, _
                                      ByVal eLetrasEspeciales As Boolean)
            If eSoloMayusculas Then
                If eLetrasEspeciales Then
                    If InStr(1, "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZÑÇ" & Chr(8), e.KeyChar) = 0 Then e.KeyChar = ""
                Else
                    If InStr(1, "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), e.KeyChar) = 0 Then e.KeyChar = ""
                End If
            Else
                If eLetrasEspeciales Then
                    If InStr(1, "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZçÇñÑ" & Chr(8), e.KeyChar) = 0 Then e.KeyChar = ""
                Else
                    If InStr(1, "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), e.KeyChar) = 0 Then e.KeyChar = ""
                End If
            End If
        End Sub
#End Region

#Region " GENERACIÓN Y VALIDACIÓN CIF"
        ''' <summary>
        ''' Indica si la cadena de texto es un cif 
        ''' </summary>
        ''' <param name="eCadena">Cadena de texto a verificar</param>
        Public Function esCif(ByVal eCadena As String) As Boolean
            Dim laExpresion As New Regex("([^A-W0-9]|[OT]|[^\w])", RegexOptions.IgnoreCase)
            eCadena = laExpresion.Replace(eCadena, String.Empty).ToUpper()

            Dim exp As New Regex("^[ABEH]\d{8}$|^[KPQS]\d{7}[A-J]$|^[CDFGJLMNRUVW]\d{7}$", RegexOptions.IgnoreCase)
            Return exp.IsMatch(eCadena.ToUpper)
        End Function

        ''' <summary>
        ''' Valida que el parámetro de entrada sea un NIF o un CIF
        ''' </summary>
        ''' <param name="eCadena">Cadena de texto a verificar</param>
        Public Function validarNIFCIF(ByVal eCadena As String) As Boolean
            Return ValidarCIF(eCadena) OrElse validarNif(eCadena)
        End Function

        ''' <summary>
        ''' Obtiene el dígito de control de un CIF
        ''' </summary>
        ''' <param name="eCIF">Cif del que se quiere obtener el dígito de control</param>
        ''' <returns>Dígito de control del CIF</returns>
        Public Function obtenerDCCIF(ByVal eCIF As String) As Char
            If eCIF Is Nothing Then Return ""

            ' Pasamos el NIF a mayúscula a la vez que eliminamos todos los
            ' carácteres que no sean números o letras. Note que no incluyo
            ' la letra I, porque si bien no puede aparecer como carácter
            ' inicial de un NIF, sí puede ser un dígito de control válido,
            ' lo que no sucede con las letras O y T.
            Dim re As New Regex("([^A-W0-9]|[OT]|[^\w])", RegexOptions.IgnoreCase)
            eCIF = re.Replace(eCIF, String.Empty).ToUpper()

            ' Para calcular el CIF, se debe de haber pasado un máximo
            ' de nueve caracteres a la función: una letra válida (que
            ' necesariamente deberá de estar comprendida en el intervalo
            ' ABCDEFGHJKLMNPQRSUVW), siete números, y el dígito de control,
            ' que puede ser un número o una letra, dependiendo de la entidad
            ' a la que pertenezca el NIF.
            '
            ' En el patrón de la expresión regular, defino dos grupos en el
            ' asiguiente orden:
            ' 1º) 1 letra + 7 u 8 dígitos.
            ' 2º) 1 letra + 7 dígitos + 1 letra.
            '
            ' Note que en el siguiente patrón, no incluyo la letra I como
            ' carácter de inicio válido del NIF.
            re = New Regex("^[ABEH]\d{8}$|^[KPQS]\d{7}[A-J]$|^[CDFGJLMNRUVW]\d{7}$", RegexOptions.IgnoreCase)
            If (Not (re.IsMatch(eCIF))) Then Return ""

            Try
                ' Guardo el último carácter pasado, siempre que
                ' el NIF disponga de nueve caracteres.
                Dim lastChar As Char = Nothing
                If (eCIF.Length = 9) Then lastChar = eCIF.Chars(8)

                Dim sumaPar, sumaImpar As Int32

                ' A continuación, la cadena debe tener 7 dígitos.
                Dim digits As String = eCIF.Substring(1, 7)
                For n As Int32 = 0 To digits.Length - 1 Step 2
                    If (n < 6) Then
                        ' Sumo las cifras pares del número que se corresponderá
                        ' con los caracteres 1, 3 y 5 de la variable «digits».
                        sumaImpar += CInt(CStr(digits.Chars(n + 1)))
                    End If

                    ' Multiplico por dos cada cifra impar (caracteres 0, 2, 4 y 6).
                    Dim dobleImpar As Int32 = 2 * CInt(CStr(digits.Chars(n)))

                    ' Acumulo la suma del doble de números impares.
                    sumaPar += (dobleImpar Mod 10) + (dobleImpar \ 10)
                Next

                ' Sumo las cifras pares e impares.
                Dim sumaTotal As Int32 = sumaPar + sumaImpar

                ' Me quedo con la cifra de las unidades y se la resto a 10, siempre
                ' y cuando la cifra de las unidades sea distinta de cero
                sumaTotal = (10 - (sumaTotal Mod 10)) Mod 10

                Dim characters() As Char = _
                            {"J"c, "A"c, "B"c, "C"c, "D"c, "E"c, "F"c, "G"c, "H"c, "I"c}

                Dim dc As Char = characters(sumaTotal)

                ' Devuelvo el Dígito de Control dependiendo del primer carácter
                ' del NIF pasado a la función.
                Dim firstChar As Char = eCIF.Chars(0)

                Select Case firstChar
                    Case "N"c, "P"c, "Q"c, "R"c, "S"c, "W"c, "K"c, "L"c, "M"c

                        ' NIF de entidades cuyo dígito de control se corresponde
                        ' con una letra. Se incluyen las letras K, L y M porque
                        ' el cálculo del dígito de control es el mismo que para
                        ' el CIF.
                        '
                        ' Al estar los índices de los arrays en base cero, el primer
                        ' elemento del array se corresponderá con la unidad del número
                        ' 10, es decir, el número cero.
                        Return characters(sumaTotal)

                    Case "C"c
                        If ((lastChar = CStr(sumaTotal)) OrElse (lastChar = dc)) Then
                            ' Devuelvo el dígito de control pasado, que
                            ' puede ser un número o una letra.
                            Return lastChar
                        Else
                            ' Devuelvo la letra correspondiente al cálculo
                            ' del dígito de control del NIF.
                            Return dc
                        End If

                    Case Else
                        ' NIF de las restantes entidades, cuyo dígito de control es un número.
                        Return CChar(CStr(sumaTotal))
                End Select
            Catch
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' Valida el CIF pasado a la función es correcto
        ''' </summary>
        ''' <param name="eCIF">CIF que se quiere validar</param>
        ''' <returns>True o False dependiendo de la validez</returns>
        Public Function validarCIF(ByRef eCIF As String) As Boolean
            Dim nifTemp As String = eCIF.Trim().ToUpper()

            If (nifTemp.Length < 9) Then Return False

            ' Guardamos el noveno carácter.
            Dim dcTemp As Char = nifTemp.Chars(8)

            ' Nos quedamos con los primeros ocho caracteres.
            nifTemp = nifTemp.Substring(0, 8)

            ' Obtengo el dígito de control correspondiente, utilizando
            ' para ello una llamada a la función GetDcCif
            '
            Dim dc As Char = obtenerDCCIF(eCIF)

            If (Not (dc = Nothing)) Then
                eCIF = nifTemp & dc
            End If

            Return (dc = dcTemp)
        End Function
#End Region

#Region " GENERACIÓN Y VALIDACIÓN NIF"
        ''' <summary>
        ''' Vefica si el parámetro de entrada es un nif válido o no
        ''' </summary>
        ''' <param name="eCadena">Cadena de texto que desemos verificar</param>
        Public Function esNif(ByVal eCadena As String) As Boolean
            Dim re As New Regex("([^A-Z0-9]|[IÑOU]|[^\w])", RegexOptions.IgnoreCase)
            eCadena = re.Replace(eCadena, String.Empty).ToUpper()

            Dim exp As New Regex("^[KLMXYZ]?\d{8}[A-HJ-NP-TV-Z]$", RegexOptions.IgnoreCase)
            Return exp.IsMatch(eCadena.ToUpper)
        End Function

        ''' <summary>
        ''' Valida el NIF introducido
        ''' </summary>
        ''' <param name="eNIF">NIF que se quiere validar</param>
        Public Function validarNif(ByRef eNIF As String) As Boolean
            If eNIF Is Nothing Then Return False
            Dim nifTemp As String = eNIF.Trim().ToUpper()

            If ((nifTemp.Length > 9) Or (nifTemp.Length = 0)) Then Return False

            ' Guardamos el dígito de control.
            Dim dcTemp As Char = nifTemp.Chars(eNIF.Length - 1)

            ' Compruebo si el dígito de control es un número.
            If (Char.IsDigit(dcTemp)) Then Return Nothing

            ' Nos quedamos con los caracteres, sin el dígito de control.
            nifTemp = nifTemp.Substring(0, eNIF.Length - 1)

            If (nifTemp.Length < 8) Then
                Dim paddingChar As String = New String("0"c, 8 - nifTemp.Length)
                nifTemp = nifTemp.Insert(nifTemp.Length, paddingChar)
            End If

            ' Obtengo el dígito de control correspondiente, utilizando
            ' para ello una llamada a la función GetDcCif.
            '
            Dim dc As Char = obtenerDCNIF(eNIF)

            If (Not (dc = Nothing)) Then
                eNIF = nifTemp & dc
            End If

            Return (dc = dcTemp)
        End Function

        ''' <summary>
        ''' Obtiene el dígito de control del NIF que se le pasa como parámtro
        ''' </summary>
        ''' <param name="eNIF">NIF del que se quiere obtener el DC</param>
        ''' <returns>Dígito de control del NIF</returns>
        Public Function obtenerDCNIF(ByVal eNIF As String) As Char
            If eNIF Is Nothing Then Return ""
            ' Pasamos el NIF a mayúscula a la vez que eliminamos los
            ' espacios en blanco al comienzo y al final de la cadena.
            eNIF = eNIF.Trim().ToUpper()

            ' El NIF está formado de uno a nueve números seguido
            ' de una letra.
            '
            ' El NIF de otros colectivos de personas físicas, está
            ' formato por una letra (K, L, M), seguido de 7 números
            ' y de una letra final.
            '
            ' El NIE está formado de una letra inicial (X, Y, Z),
            ' seguido de 7 números y de una letra final.
            ' 
            ' En el patrón de la expresión regular, defino cuatro grupos en el
            ' asiguiente orden:
            '
            ' 1º) 1 a 8 dígitos.
            ' 2º) 1 a 8 dígitos + 1 letra.
            ' 3º) 1 letra + 1 a 7 dígitos 
            ' 4º) 1 letra + 1 a 7 dígitos + 1 letra.
            Try
                Dim re As New Regex("(^\d{1,8}$)|(^\d{1,8}[A-Z]$)|(^[K-MX-Z]\d{1,7}$)|(^[K-MX-Z]\d{1,7}[A-Z]$)", RegexOptions.IgnoreCase)

                If (Not (re.IsMatch(eNIF))) Then Return Nothing

                ' Nos quedamos únicamente con los números del NIF, y
                ' los formateamos con ceros a la izquierda si su
                ' longitud es inferior a siete caracteres.
                re = New Regex("(\d{1,8})")

                Dim numeros As String = re.Match(eNIF).Value.PadLeft(7, "0"c)

                ' Primer carácter del NIF.
                Dim firstChar As Char = eNIF.Chars(0)

                ' Si procede, reemplazamos la letra del NIE
                ' por el peso que le corresponde.
                If (firstChar = "X"c) Then
                    numeros = "0" & numeros
                ElseIf (firstChar = "Y"c) Then
                    numeros = "1" & numeros
                ElseIf (firstChar = "Z"c) Then
                    numeros = "2" & numeros
                End If

                ' Tabla del NIF
                '
                '  0T  1R  2W  3A  4G  5M  6Y  7F  8P  9D
                ' 10X 11B 12N 13J 14Z 15S 16Q 17V 18H 19L
                ' 20C 21K 22E 23T
                '
                ' Procedo a calcular el NIF/NIE
                Dim dni As Integer = CInt(numeros)

                ' La operación consiste en calcular el resto de dividir el DNI
                ' entre 23 (sin decimales). Dicho resto (que estará entre 0 y 22),
                ' se busca en la tabla y nos da la letra del NIF.
                '
                ' Obtenemos el resto de la división.
                Dim r As Integer = dni Mod 23

                ' Obtenemos el dígito de control del NIF
                Dim dc As Char = CChar("TRWAGMYFPDXBNJZSQVHLCKE".Substring(r, 1))

                Return dc
            Catch
                ' Cualquier excepción producida, devolverá el valor Nothing.
                Return Nothing
            End Try
        End Function

#Region " OBTENCIÓN Y VALIDACIÓN DC CCC"
        ''' <summary>
        ''' Obtiene el dígito de control de un número de cuenta bancaria
        ''' </summary>
        ''' <param name="eNumeroCuenta">El número de cuenta sin el DC</param>
        ''' <returns>Dígito de control de la cuenta bancaria</returns>
        Public Function obtenerDCCuentaBancaria(ByVal eNumeroCuenta As String) As String
            ' Primero compruebo que la longitud del parámetro
            ' sea de 18 caracteres, y que éstos sean números.
            If eNumeroCuenta.Length <> 18 Then
                Return ""
            Else
                Dim ch As Char
                For Each ch In eNumeroCuenta
                    If Not Char.IsNumber(ch) Then Return ""
                Next
            End If

            Dim cociente1, cociente2, resto1, resto2 As Integer
            Dim sucursal, cuenta, dc1, dc2 As String
            Dim suma1, suma2, n As Integer

            ' Matriz que contiene los pesos utilizados en el
            ' algoritmo de cálculo de los dos dígitos de control.
            Dim pesos() As Integer = {6, 3, 7, 9, 10, 5, 8, 4, 2, 1}

            sucursal = eNumeroCuenta.Substring(0, 8)
            cuenta = eNumeroCuenta.Substring(8, 10)

            ' Obtengo el primer dígito de control que verificará
            ' los códigos de Entidad y Oficina.
            For n = 7 To 0 Step -1
                suma1 = suma1 + Convert.ToInt32(sucursal.Substring(n, 1)) * pesos(7 - n)
            Next

            ' Calculamos el cociente de dividir el resultado
            ' de la suma entre 11.
            cociente1 = suma1 \ 11 ' Nos da un resultado entero.

            ' Calculo el resto de la división. Para ello,
            ' en lugar de utilizar el operador Mod, utilizo
            ' la fórmula para obtener el resto de la división.
            resto1 = suma1 - (11 * cociente1)

            dc1 = (11 - resto1).ToString

            Select Case dc1
                Case "11"
                    dc1 = "0"
                Case "10"
                    dc1 = "1"
            End Select

            ' Ahora obtengo el segundo dígíto, que verificará
            ' el número de cuenta de cliente.
            For n = 9 To 0 Step -1
                suma2 = suma2 + Convert.ToInt32(cuenta.Substring(n, 1)) * pesos(9 - n)
            Next

            ' Calculamos el cociente de dividir el resultado
            ' de la suma entre 11.
            cociente2 = suma2 \ 11 ' Nos da un resultado entero

            ' Calculo el resto de la división. Para ello,
            ' en lugar de utilizar el operador Mod, utilizo
            ' la fórmula para obtener el resto de la división.
            resto2 = suma2 - (11 * cociente2)

            dc2 = (11 - resto2).ToString

            Select Case dc2
                Case "11"
                    dc2 = "0"
                Case "10"
                    dc2 = "1"
            End Select

            ' Devuelvo el dígito de control.
            Return dc1 & dc2
        End Function


        ''' <summary>
        ''' Valída el dígito de control de un número de cuenta bancaria
        ''' </summary>
        ''' <param name="eNumeroCuenta">Número de cuenta que se quiere verificar</param>
        ''' <returns>True o False si la cuenta es correcta</returns>
        Public Function ValidarCuentaBancaria(ByVal eNumeroCuenta As String) As Boolean
            ' Primero compruebo que la longitud del parámetro
            ' sea de 20 caracteres.
            If eNumeroCuenta.Length <> 20 Then Return False

            ' Extraigo el dígito de control.
            Dim dc As String = eNumeroCuenta.Substring(8, 2)

            ' Del número de cuenta, elimino el dígito de control.
            eNumeroCuenta = eNumeroCuenta.Remove(8, 2)

            ' Obtengo el dígito de control verdadero.
            Dim dcTemp As String = obtenerDCCuentaBancaria(eNumeroCuenta)

            ' Devuelvo el resultado.
            Return (dc = dcTemp)
        End Function
#End Region

#Region " OBTENCIÓN VALIDACIÓN DC EAN13 "
        ''' <summary>
        ''' Calcula el DC de un código de barras en formato EAN13
        ''' </summary>
        ''' <param name="eNumero">Obtiene el número de conterol de un código EAN13</param>
        ''' <returns>Número de control</returns>
        Public Function ObtenerDCEAN13(ByRef eNumero As String) As Integer
            ' Si el número no cumple con el formato EAN13, la función
            ' devolverá una excepción de argumentos no válidos. 
            ' Para ello, la cadena deberá tener 12 caracteres de
            ' longitud, y contener sólo números.
            If eNumero.Length <> 12 Then
                eNumero = ""
                Throw New System.ArgumentException
            Else
                Dim ch As Char
                For Each ch In eNumero
                    If Not Char.IsNumber(ch) Then
                        eNumero = ""
                        Throw New System.ArgumentException
                    End If
                Next
            End If

            Dim x, digito, sumaCod As Integer

            ' Extraigo el valor del dígito, y voy
            ' sumando los valores resultantes.
            For x = 11 To 0 Step -1

                digito = CInt(eNumero.Substring(x, 1))

                Select Case x
                    Case 1, 3, 5, 7, 9, 11
                        ' Las posiciones impares se multiplican por 3
                        sumaCod += (digito * 3)

                    Case Else
                        ' Las posiciones pares se multiplican por 1
                        sumaCod += (digito * 1)
                End Select
            Next

            ' Calculo la decena superior
            digito = (sumaCod Mod 10)

            ' Calculo el dígito de control
            digito = 10 - digito

            ' Código de barras completo
            eNumero &= CStr(digito)

            ' Devuelvo el dígito de control
            Return digito
        End Function

        ''' <summary>
        ''' Valída el DC de un EAN13
        ''' </summary>
        ''' <param name="eNumero">Valida un EAN13</param>
        ''' <returns>True o False dependiendo de si es correcto</returns>
        Public Function ValidarEAN13(ByVal eNumero As String) As Boolean
            ' Si el número no cumple con el formato EAN13, 
            ' la función devolverá una excepción de argumentos
            ' no válidos.
            If eNumero.Length <> 13 Then
                Throw New System.ArgumentException
            Else
                Dim ch As Char
                For Each ch In eNumero
                    If Not Char.IsNumber(ch) Then
                        Throw New System.ArgumentException
                    End If
                Next
            End If

            Dim digito, lastDigit As Integer

            ' Último dígito del número.
            lastDigit = CInt(eNumero.Substring(eNumero.Length - 1, 1))

            ' Calculo el dígito de control del número pasado
            digito = ObtenerDCEAN13(eNumero.Substring(0, 12))

            ' Compruebo si los dos dígitos son iguales
            Return (digito = lastDigit)
        End Function
#End Region

#Region " VALIDACIÓN DE TARJETAS DE CRÉDITO "
        ''' <summary>
        ''' Valída la numeración de una Tarjeta de Crédito para comprobar si es real
        ''' </summary>
        ''' <param name="eNumeroTarjeta">Número de tarjeta que se quiere comprobar</param>
        ''' <returns>True o False dependiendo de si la tarjeta es válida</returns>
        Public Function validarTarjetaCredito(ByVal eNumeroTarjeta As String) As Boolean
            '*******************************************************************
            ' El algoritmo ISO 2894 consiste en lo siguiente:
            '
            ' 1. Calcular el peso para el primer dígito: si el número de
            '    dígitos es par el primer peso es 2, de lo contrario es 1.
            '    Después los pesos alternan entre 1, 2, 1, 2, 1 ...
            ' 2. Multiplicar cada dígito por su peso.
            ' 3. Si el resultado del 2º peso es mayor que 9, restar 9.
            ' 4. Sumar todos los dígitos.
            ' 5. Comprobar que el resultado es divisible por 10.
            '*******************************************************************

            Dim peso, digito, suma, x As Integer
            Dim newCard As String = ""
            Dim ch As Char

            ' Elimino de la cadena cualquier carácter que no sea un número.
            For Each ch In eNumeroTarjeta
                If Char.IsNumber(ch) Then
                    newCard = newCard & ch
                End If
            Next

            ' Si es 0 devolver una excepción de argumentos no permitidos
            If newCard.Length = 0 Then Throw New System.ArgumentException

            ' Si el número de dígitos es par, el primer peso es 2,
            ' de lo contrario es 1.
            peso = CInt(IIf((newCard.Length Mod 2) = 0, 2, 1))

            For x = 0 To newCard.Length - 1
                ' Extraigo los dígitos
                digito = CInt(newCard.Substring(x, 1)) * peso

                If digito > 9 Then digito += -9

                suma += digito

                ' Cambiar el peso para el siguiente dígito
                peso = CInt(IIf(peso = 2, 1, 2))
            Next

            ' Devolver verdadero si la suma es divisible por 10                
            Return ((suma Mod 10) = 0)
        End Function
#End Region

#Region " OBTENCIÓN VALIDACIÓN DEL NÚMERO DE LA SS "
        ''' <summary>
        ''' Devuelve el dígito de control de un Número de Afiliación como de
        ''' un Número de Cuenta de Cotización.
        ''' </summary>
        ''' <param name="eNumSegSocial">Número cuyo dígito de control se desea obtener.</param>
        ''' <param name="eEsNumEmpresa">Indica si el Número se corresponde con el de una Empresa.</param>
        ''' <returns>Dígito de control del número de la SS</returns>
        Private Function ObtenerDCNumSegSocial(ByVal eNumSegSocial As String, _
                                               ByVal eEsNumEmpresa As Boolean) As String
            ' Si hay más de 10 dígitos en el número se devolverá una excepción de
            ' argumentos no permitidos.
            If (eNumSegSocial.Length > 10) OrElse (eNumSegSocial.Length = 0) Then _
                Throw New System.ArgumentException()

            ' Si algún carácter no es un número, abandono la función.
            Dim regex As New System.Text.RegularExpressions.Regex("[^0-9]")
            If (regex.IsMatch(eNumSegSocial)) Then _
                Throw New System.ArgumentException()

            Try
                ' Obtengo el número correspondiente a la Provincia
                Dim dcProv As String = eNumSegSocial.Substring(0, 2)

                ' Obtengo el resto del número
                Dim numero As String = eNumSegSocial.Substring(2, eNumSegSocial.Length - 2)

                Select Case numero.Length
                    Case 8
                        If (eEsNumEmpresa) Then
                            ' Si el número es de una empresa, no puede tener 8 dígitos.
                            Return String.Empty
                        Else
                            ' Compruebo si es un NAF nuevo o antiguo.
                            If (numero.Chars(0) = "0"c) Then
                                ' Es un número de afiliación antiguo. Lo formateo
                                ' a siete dígitos, eliminando el primer cero.
                                numero = numero.Remove(0, 1)
                            End If
                        End If

                    Case 7
                        ' Puede ser un NAF antiguo o un CCC nuevo o viejo.
                        If (eEsNumEmpresa) Then
                            ' Si el primer dígito es un cero, es un CCC antiguo,
                            ' por lo que lo formateo a seis dígitos, eliminando
                            ' el primer cero.
                            If (numero.Chars(0) = "0"c) Then
                                numero = numero.Remove(0, 1)
                            End If
                        End If

                    Case 6
                        ' Si se trata del número de una empresa,
                        ' es un CCC antiguo.
                        If (Not (eEsNumEmpresa)) Then
                            ' Es un NAF antiguo, por lo que lo debo
                            ' de formatear a 7 dígitos.
                            numero = numero.PadLeft(7, "0"c)
                        End If

                    Case Else
                        ' Todos los demás casos, serán números antiguos
                        If (eEsNumEmpresa) Then
                            ' Lo formateo a seis dígitos.
                            numero = numero.PadLeft(6, "0"c)
                        Else
                            ' Lo formateo a siete dígitos.
                            numero = numero.PadLeft(7, "0"c)
                        End If
                End Select

                ' Completo el número de Seguridad Social
                Dim naf As Int64 = Convert.ToInt64(dcProv & numero)

                ' Calculo el Dígito de Control. Tengo que operar con números
                ' Long, para evitar el error de desbordamiento que se puede
                ' producir con los nuevos números de Seguridad Social
                naf = naf - (naf \ 97) * 97

                ' Devuelvo el Dígito de Control formateado
                Return String.Format("{0:00}", naf)
            Catch
                Return String.Empty
            End Try
        End Function

        ''' <summary>
        ''' Obtiene dígitos de control de números de Afiliación
        ''' </summary>
        ''' <param name="eNumSegSocial">Número d la seguridad social</param>
        ''' <returns>Obtiene dígitos de control de números de Afiliación</returns>
        Public Function obtenerDCNumSegSocial(ByVal eNumSegSocial As String) As String
            Return obtenerDCNumSegSocial(eNumSegSocial, False)
        End Function

        ''' <summary>
        ''' Obtiene dígitos de control de números de Empresa
        ''' </summary>
        ''' <param name="eNumSegSocial">Número d la seguridad social</param>
        ''' <returns>Obtiene dígitos de control de números de Empresa</returns>
        ''' <remarks></remarks>
        Public Function ObtenerDCNumEmpresa(ByVal eNumSegSocial As String) As String
            Return ObtenerDCNumSegSocial(eNumSegSocial, True)
        End Function

        ''' <summary>
        ''' Valida el Número de Afiliación o de Cuenta de Cotización.
        ''' </summary>
        ''' <param name="eNumSegSocial">Número cuyo dígito de control se desea validar.</param>
        ''' <param name="eEsNumEmpresa">Indica si el Número se corresponde con el de una Empresa.</param>
        ''' <returns>True o False dependiendo de si la validación se cumple</returns>
        ''' <remarks></remarks>
        Public Function ValidarNumSegSocial(ByRef eNumSegSocial As String, _
                                            ByVal eEsNumEmpresa As Boolean) As Boolean
            ' Si se ha pasado una cadena de longitud cero, o
            ' la longitud es superior a 12 caracteres, la función
            ' devolverá una excepción de argumentos no permitidos.
            If (eNumSegSocial.Length > 12) OrElse (eNumSegSocial.Length = 0) Then
                eNumSegSocial = String.Empty
                Throw New System.ArgumentException()
            End If

            ' Si algún carácter no es un número, abandono la función.
            '
            Dim regex As New System.Text.RegularExpressions.Regex("[^0-9]")
            If (regex.IsMatch(eNumSegSocial)) Then
                eNumSegSocial = String.Empty
                Throw New System.ArgumentException()
            End If

            Try
                ' El número de Seguridad Social se compone de 11 ó 12
                ' números, dependiendo que sea un Código de Cuenta de 
                ' Cotización o un Número de Afiliación respectivamente.
                '
                ' A su vez, el número se divide en tres partes:
                ' 1ª) Al Código Provincial corresponden las dos primeras cifras
                Dim cp As String = eNumSegSocial.Substring(0, 2)

                ' Formateo el código provincial. NOTA: Se asume que el
                ' número empezará por 0, para aquellas provincias cuyo
                ' código provincial se encuentre entre 1 y 9.
                cp = String.Format("{0:00}", CInt(cp))

                ' 2ª) Las dos últimas cifras se corresponden con el Dígito de Control
                Dim dcTemp As String = eNumSegSocial.Substring(eNumSegSocial.Length - 2)
                dcTemp = String.Format("{0:00}", CInt(dcTemp))

                ' 3ª) Las cifras intermedias, entre el DP y el DC, se corresponden
                ' con el número propiamente dicho.
                Dim numero As String = eNumSegSocial.Substring(2, eNumSegSocial.Length - 4)

                ' Formateamos el número, dependiendo de que sea
                ' un número de afiliación o un número de empresa
                If (eEsNumEmpresa) Then
                    ' 11 dígitos    --> Es el Código de Cuenta de Cotización
                    numero = String.Format("{0:0000000}", Convert.ToInt64(numero))
                Else
                    ' 12 dígitos    --> Es el Número de Afiliación
                    numero = String.Format("{0:00000000}", Convert.ToInt64(numero))
                End If

                ' Obtenemos el número de Seguridad Social sin DC
                numero = cp & numero

                ' Calculo el dígito de control verdadero haciendo
                ' una llamada a la función GetDCSegSocial.
                Dim dc As String = ObtenerDCNumSegSocial(numero, eEsNumEmpresa)
                If (dc = String.Empty) Then
                    eNumSegSocial = String.Empty
                    Return False
                End If

                ' Número de Seguridad Social resultante.
                eNumSegSocial = numero & dc

                ' Comparo los dos dígitos de control
                If (dc.Equals(dcTemp)) Then
                    ' Es correcto.
                    Return True
                Else
                    ' No es correcto. 
                    Return False
                End If

            Catch
                ' Cualquier excepción producida, devolverá False.                    
                eNumSegSocial = String.Empty
                Return False
            End Try
        End Function
#End Region
#End Region
    End Module
End Namespace
