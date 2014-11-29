Imports System.Windows.Forms

Namespace WinForms
    Namespace ProgressBar
        Public Module modRecompilaWinFormsProgressBar
            ''' <summary>
            ''' Fija el valor de la barra de progreso
            ''' </summary>
            ''' <param name="eBarra">Barra de progreso a la que se le va a asignar el valor</param>
            ''' <param name="eValor">Valor que se va a fijar en la barra de progreso</param>
            Public Sub FijarBarra(ByRef eBarra As Object, _
                                  ByVal eValor As Long)
                Try
                    If eValor < eBarra.Maximum Then
                        eBarra.Value = eValor
                    Else
                        eBarra.Value = eBarra.Maximum
                    End If
                    Application.DoEvents()
                Catch ex As Exception
                End Try
            End Sub

            ''' <summary>
            ''' Aumenta el valor de la barra de progreso, si no se especifica ningún valor, este será
            ''' tomado como en 1
            ''' </summary>
            ''' <param name="eBarra">Barra de progreso a la que se le va a aumenter el valor</param>
            ''' <param name="eValor">Valor a aumentar en la barra de progreso</param>
            Public Sub AumentarBarra(ByRef eBarra As Object, _
                                     Optional ByVal eValor As Long = 1)
                Try
                    If eBarra.Value + eValor < eBarra.Maximum Then
                        eBarra.Value += eValor
                    Else
                        eBarra.Value = eBarra.Maximum
                    End If
                    Application.DoEvents()
                Catch ex As Exception
                End Try
            End Sub

            ''' <summary>
            ''' Fija el valor máximo que puede tomar una barra de progreso
            ''' </summary>
            ''' <param name="eBarra">Barra de progreso a la que se le va a fijar el valor máximo</param>
            ''' <param name="eMaximo">Valor máximo que puede tomar la barra</param>
            Public Sub fijarMaximoBarra(ByRef eBarra As Object, _
                                        Optional ByVal eMaximo As Long = 100)
                Try
                    eBarra.Maximum = eMaximo
                    Application.DoEvents()
                Catch ex As Exception
                End Try
            End Sub

            ''' <summary>
            ''' Fija el valor mínimo que puede tomar una barra de progreso
            ''' </summary>
            ''' <param name="eBarra">Barra de progreso a la que se le va a asignar el mínimo valor</param>
            ''' <param name="eMinimo">Mínimo valor que puede tomar la barra</param>
            ''' <remarks></remarks>
            Public Sub fijarMinimoBarra(ByRef eBarra As Object, _
                                        Optional ByVal eMinimo As Long = 0)
                Try
                    eBarra.Minimum = eMinimo
                    Application.DoEvents()
                Catch ex As Exception
                End Try
            End Sub
        End Module
    End Namespace
End Namespace

