Namespace Funciones
    Public Module modRecompilaFunciones
        ''' <summary>
        ''' Convierte un Boolean en un Bit
        ''' </summary>
        ''' <param name="elBoolean">Valor boolean</param>
        ''' <returns>1 = True | 0 = False</returns>
        Public Function booleanABit(ByVal elBoolean As Boolean) As Integer
            If elBoolean = True Then
                Return 1
            Else
                Return 0
            End If
        End Function

        ''' <summary>
        ''' Convierte un Bit en Boolean
        ''' </summary>
        ''' <param name="elBit">Balor bit</param>
        ''' <returns>True = 1 | False = 0</returns>
        Public Function bitABoolean(ByVal elBit As String) As Boolean
            Return (Funciones.NZI(elBit) = 1)
        End Function

        ''' <summary>
        ''' Convierte un texto Si o No a boolean
        ''' </summary>
        ''' <param name="laCadena">Texto Si o No</param>
        ''' <returns>Si = True | No = False</returns>
        Public Function SiONoaBoolean(ByVal laCadena As String) As Boolean
            Return (laCadena.Trim.ToUpper = "SI")
        End Function

        ''' <summary>
        ''' Convierte un valor booleano a Si o No
        ''' </summary>
        ''' <param name="elBooleano">Valor booleano</param>
        ''' <returns>True = Si | False = No</returns>
        Public Function booleanASiONo(ByVal elBooleano As Boolean) As String
            If elBooleano Then
                Return "Si"
            Else
                Return "No"
            End If
        End Function

        Public Function SiONoaBit(ByVal laCadena As String) As Integer
            Return booleanABit(SiONoaBoolean(laCadena))
        End Function

        Public Function bitASioNO(ByVal elBit As Integer) As String
            Return booleanASiONo(bitABoolean(elBit))
        End Function

        Public Function IntABinario(ByVal eInt As Long) As String
            Return Convert.ToString(eInt, 2)
        End Function

        Public Function IntAHexadecimal(ByVal eInt As Long) As String
            Return Convert.ToString(eInt, 16)
        End Function

        Public Function BinarioAInt(ByVal eBinario As String) As Long
            Return Convert.ToInt32(eBinario, 2)
        End Function

        Public Function HexadecimalAInt(ByVal eHex As Long) As String
            Return Convert.ToInt32(eHex, 16)
        End Function

        Public Function NZL(ByVal laCadena As String) As Long
            If IsNumeric(laCadena) Then
                Return CLng(laCadena)
            Else
                Return 0
            End If
        End Function

        Public Function NZI(ByVal laCadena As String) As Integer
            If IsNumeric(laCadena) Then
                Return CInt(laCadena)
            Else
                Return 0
            End If
        End Function

        Public Function NZSh(ByVal laCadena As String) As Short
            If IsNumeric(laCadena) Then
                Return CShort(laCadena)
            Else
                Return 0
            End If
        End Function

        Public Function NZD(ByVal laCadena As String) As Double
            If IsNumeric(laCadena) Then
                Return CDbl(laCadena)
            Else
                Return 0
            End If
        End Function

        Public Function NZS(ByVal laCadena As String) As Single
            If IsNumeric(laCadena) Then
                Return CSng(laCadena)
            Else
                Return 0
            End If
        End Function

        Public Function QuitarComillas(ByVal cadena As String) As String
            Dim cadena2 As String
            cadena2 = cadena
            cadena2 = Replace(cadena2 & "", """", "")
            cadena2 = Replace(cadena2 & "", "'", "")
            Return (cadena2 & "")
        End Function

        ''' <summary>
        ''' Convierte una cadena en un decimal, utilizando la última cóma o punto
        ''' como separador decimal
        ''' </summary>
        ''' <param name="eCadena">Cadena que se va a convertir</param>
        ''' <returns>Dobule generado a partir de la cadena</returns>
        Public Function string2decimal(ByVal eCadena As String) As Double
            Dim paraDevolver As Double = 0.0

            ' Se cambian todas las comas por puntos y después se
            ' hace un split de la cadena por los puntos. La última parte
            ' se considera la parte decimal, mientras que el resto se considera
            ' la parte entera
            eCadena = eCadena.Replace(",", ".")
            If Not eCadena.Contains(".") Then eCadena &= ".00"
            Dim lasPartes() As String = eCadena.Split(".")
            If lasPartes.Length > 0 Then
                Dim parteDecimal As String = lasPartes(lasPartes.Length - 1)
                Dim parteEntera As String = ""
                For i As Integer = 0 To lasPartes.Length - 2
                    parteEntera &= lasPartes(i)
                Next
                Dim elResultado As String = parteEntera & "," & parteDecimal
                Double.TryParse(elResultado, paraDevolver)
            End If
            Return paraDevolver
        End Function

        Function GenerarIDUnico(Optional ByVal _restoCadena As String = "") As String
            Dim laFecha As Date = Now
            Dim paraDevolver As String = ""
            Dim i As Integer


            i = laFecha.Year - 2008
            If i < 0 Or i > 35 Then i = 0
            If i < 10 Then
                paraDevolver = Chr(Asc("0") + i)
            Else
                paraDevolver = Chr(Asc("A") + i - 10)
            End If

            i = laFecha.Month
            If i < 10 Then
                paraDevolver &= Chr(Asc("0") + i)
            Else
                paraDevolver &= Chr(Asc("A") + i - 10)
            End If

            i = laFecha.Day
            If i < 10 Then
                paraDevolver &= Chr(Asc("0") + i)
            Else
                paraDevolver &= Chr(Asc("A") + i - 10)
            End If

            i = laFecha.Hour
            If i < 10 Then
                paraDevolver &= Chr(Asc("0") + i)
            Else
                paraDevolver &= Chr(Asc("A") + i - 10)
            End If

            i = laFecha.Minute
            paraDevolver &= Format$(i, "00")

            i = laFecha.Second
            paraDevolver &= Format$(i, "00")

            paraDevolver &= _restoCadena

            Return (paraDevolver)
        End Function
    End Module
End Namespace
