Namespace Aleatorios
    ''' <summary>
    ''' Funciones útilies para la generación de números y cadenas aleatorias
    ''' </summary>
    Public Module modAleatorios
#Region " PROPIEDADES "
        ''' <summary>
        ''' Semilla para la generación de números aleatorios.
        ''' Esta semilla se inicializa la primera vez que se utiliza para 
        ''' facilitar la generación de números aleatorios.
        ''' </summary>
        Public ReadOnly Property Semilla As Integer
            Get
                If iSemilla = -1 Then
                    iSemilla = DateTime.Now.Millisecond
                End If
                Return iSemilla
            End Get
        End Property
        Private iSemilla As Integer = -1

        ''' <summary>
        ''' Objeto Random que se encargará de la generación de los 
        ''' números aleatorios.
        ''' </summary>
        Public ReadOnly Property Generador As Random
            Get
                If iGenerador Is Nothing Then
                    iGenerador = New Random(Semilla)
                End If
                Return iGenerador
            End Get
        End Property
        Private iGenerador As Random = Nothing
#End Region

#Region " METODOS "
        ''' <summary>
        ''' Genera un número entero aleatorio entre unn rango preestablecido
        ''' </summary>
        ''' <param name="eInicio">Número mínimo que se puede obtener de forma aleatoria</param>
        ''' <param name="eFin">Número máximo que se puede obtener de forma aleatorioa</param>
        ''' <returns>Un número aleatorio entre el mínimo y el máximo pasados como parámetro</returns>
        Public Function enteroAleatorio(ByVal eInicio As Integer, _
                                        ByVal eFin As Integer) As Integer
            Return Generador.Next(eInicio, eFin)
        End Function


        ''' <summary>
        ''' Genera una cadena de texto aleatoria
        ''' </summary>
        ''' <param name="eNumeroCaracteres">Número de caracteres que tiene que tener la cadena</param>
        ''' <param name="eSoloAlfanumericos">Sólo utiliza caracteres alfanuméricos para la generación</param>
        ''' <returns>Cadena aleatoria con la longitud especificada</returns>
        Public Function cadenaAleatoria(ByVal eNumeroCaracteres As Integer, _
                                        Optional ByVal eSoloAlfanumericos As Boolean = False) As String
            Dim paraDevolver As String = String.Empty

            Dim caracterInicio As Integer = IIf(eSoloAlfanumericos, 48, 33)
            Dim caracterFin As Integer = IIf(eSoloAlfanumericos, 122, 126)
            Dim aux As Integer = 0

            For i As Integer = 0 To eNumeroCaracteres - 1
                Do
                    aux = enteroAleatorio(caracterInicio, caracterFin)
                Loop While (eSoloAlfanumericos AndAlso (((aux > 57) And (aux < 65)) OrElse ((aux > 90) And (aux < 97))))
                paraDevolver &= Chr(aux)
            Next

            Return paraDevolver
        End Function

        ''' <summary>
        ''' Genera una cadena aleatoria de una determinada longitud utilizando solamente
        ''' caracteres hexadecimales
        ''' </summary>
        ''' <param name="eNumeroCaracteres">Número de caracteres que tiene que tener la cadena</param>
        ''' <returns>Cadena hexadecimal aleatoria con la longitud especificada</returns>
        Public Function cadenaAleatoriaHex(ByVal eNumeroCaracteres As Integer) As String
            Dim paraDevolver As String = String.Empty

            Dim caracterInicio As Integer = 48
            Dim caracterFin As Integer = 71
            Dim aux As Integer = 0

            For i As Integer = 0 To eNumeroCaracteres - 1
                Do
                    aux = enteroAleatorio(caracterInicio, caracterFin)
                Loop While ((aux > 57) And (aux < 65))
                paraDevolver &= Chr(aux)
            Next

            Return paraDevolver
        End Function
#End Region
    End Module
End Namespace