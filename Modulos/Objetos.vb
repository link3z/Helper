Imports System.Reflection

Namespace Objetos
    ''' <summary>
    ''' Funciones útiles para trabajar con objetos 
    ''' </summary>
    ''' <remarks></remarks>
    Public Module modObjetos
        ''' <summary>
        ''' Comprueba si un objeto tiene una propiedad especificada
        ''' </summary>
        ''' <param name="eObjeto">Objeto que se quiere comprobar</param>
        ''' <param name="ePropiedad">Nombre de la propiedad que se tiene que comprobar</param>
        ''' <returns>True | False dependiendo de si existe o no la propiedad solicitada</returns>   
        ''' <remarks>Se debe especificar un objeto y un nombre para la propiedad,
        ''' en caso contrario se lanzará una excepción</remarks>         
        Public Function tienePropiedad(ByVal eObjeto As Object,
                                       ByVal ePropiedad As String) As Boolean
            ' Se verifican que los parámetros que se le pasaron a la función sean
            ' los correctos, en caso contrario se lanzan excepciones.
            If eObjeto Is Nothing Then
                Throw New Exception("El objeto a comprobar no puede ser Nothing.")
            End If

            If String.IsNullOrEmpty(ePropiedad) Then
                Throw New Exception("No se ha especificado el nombre de la propiedad a comprobar.")
            End If

            Dim type As Type = eObjeto.GetType
            Return type.GetProperty(ePropiedad) IsNot Nothing
        End Function

        ''' <summary>
        ''' Comprueba si un objeto tiene un metodo en especificado
        ''' </summary>
        ''' <param name="eObjeto">Objeto que se quiere comprobar</param>
        ''' <param name="eMetodo">Nombre del método a ser comprobado</param>
        ''' <returns>True | False dependiendo de si existe o no el método solicitado</returns>          
        Public Function tieneMetodo(ByVal eObjeto As Object, _
                                    ByVal eMetodo As String) As Boolean
            ' Se verifican que los parámetros que se le pasaron a la función sean
            ' los correctos, en caso contrario se lanzan excepciones.
            If eObjeto Is Nothing Then
                Throw New Exception("El objeto a comprobar no puede ser Nothing.")
            End If

            If String.IsNullOrEmpty(eMetodo) Then
                Throw New Exception("No se ha especificado el nombre del método a comprobar.")
            End If

            Dim type As Type = eObjeto.GetType
            Return type.GetMethod(eMetodo) IsNot Nothing
        End Function

        ''' <summary>
        ''' Obtiene el valor de una propiedad de un objeto a partir del objeto y el nombre de la proiedad
        ''' </summary>
        ''' <param name="eObjeto">Objeto del que se quiere obtener el valor de la propiedad</param>
        ''' <param name="eNombrePropiedad">Nombre de la propiedad</param>
        ''' <param name="eConExcepcion">Lanza excepción si el objeto no tiene la propiedad</param>
        ''' <returns>El valor de la propiedad del objeto</returns>
        ''' <remarks>Se debe especificar un objeto y un nombre para la propiedad,
        ''' en caso contrario se lanzará una excepción</remarks>  
        Public Function obtenerValorPropiedad(ByVal eObjeto As Object, _
                                              ByVal eNombrePropiedad As String, _
                                              Optional ByVal eConExcepcion As Boolean = True) As Object
            If tienePropiedad(eObjeto, eNombrePropiedad) Then
                Return eObjeto.[GetType]().GetProperty(eNombrePropiedad).GetValue(eObjeto, Nothing)
            Else
                If eConExcepcion Then
                    Throw New Exception("El objeto '" & eObjeto.ToString & "' no tiene la propiedad '" & eNombrePropiedad & "'.")
                Else
                    Return Nothing
                End If
            End If
        End Function

        ''' <summary>
        ''' Asigna el valor a una propiedad de un objeto a partir del objeto y el nombre de la proiedad
        ''' </summary>
        ''' <param name="eObjeto">Objeto del que se quiere asignar el valor de la propiedad</param>
        ''' <param name="eNombrePropiedad">Nombre de la propiedad</param>
        ''' <param name="eValor">Valor que se le va a asignar a la propiedad</param>
        ''' <param name="eConExcepcion">Lanza excepción si el objeto no tiene la propiedad</param>
        ''' <returns>True | False dependiendo de si puedo realizar la operación correctamente</returns>   
        ''' <remarks>Se debe especificar un objeto y un nombre para la propiedad,
        ''' en caso contrario se lanzará una excepción</remarks>  
        Public Function asignarValorPropiedad(ByVal eObjeto As Object, _
                                              ByVal eNombrePropiedad As String, _
                                              ByVal eValor As Object, _
                                              Optional ByVal eConExcepcion As Boolean = False) As Boolean
            If tienePropiedad(eObjeto, eNombrePropiedad) Then
                Try
                    Dim infoPropiedad As System.Reflection.PropertyInfo = eObjeto.GetType().GetProperty(eNombrePropiedad)
                    eValor = Convert.ChangeType(eValor, infoPropiedad.PropertyType)
                    infoPropiedad.SetValue(eObjeto, eValor, Nothing)

                    Return True
                Catch ex As Exception
                    If eConExcepcion Then
                        Throw New Exception("Se ha producido un error al tratar de asignar el valor a la propiedad '" & eNombrePropiedad & "'.", ex)
                    Else
                        Return False
                    End If
                End Try
            Else
                If eConExcepcion Then
                    Throw New Exception("El objeto '" & eObjeto.ToString & "' no tiene la propiedad '" & eNombrePropiedad & "'.")
                Else
                    Return Nothing
                End If
            End If
        End Function

        ''' <summary>
        ''' Obtiene el valor de un méteodo de un objeto a partir del objeto y el nombre del método
        ''' </summary>
        ''' <param name="eObjeto">Objeto del que se quiere obtener el valor del método</param>
        ''' <param name="eNombreMetodo">Nombre del método</param>
        ''' <param name="eParametros">Si el método necesita parámetros para la invocación, array con los
        ''' parámetros necesarios para la invocación</param>
        ''' <param name="eConExcepcion">Lanza excepción si el objeto no tiene el método o la invocación
        ''' no es correcta</param>
        ''' <returns>True | False dependiendo de si puedo realizar la operación correctamente</returns>  
        ''' <remarks>Se debe especificar un objeto y un nombre para la propiedad,
        ''' en caso contrario se lanzará una excepción</remarks>  
        Public Function obtenerValorMetodo(ByVal eObjeto As Object, _
                                           ByVal eNombreMetodo As String,
                                           Optional ByVal eParametros() As Object = Nothing, _
                                           Optional ByVal eConExcepcion As Boolean = True) As Object
            If tieneMetodo(eObjeto, eNombreMetodo) Then
                Try
                    Return eObjeto.[GetType]().GetMethod(eNombreMetodo).Invoke(eObjeto, eParametros)
                Catch ex As Exception
                    If eConExcepcion Then
                        Throw New Exception("Se ha producido un error al tratar de invocar el método '" & eNombreMetodo & "'.", ex)
                    Else
                        Return Nothing
                    End If
                End Try
            Else
                If eConExcepcion Then
                    Throw New Exception("El objeto '" & eObjeto.ToString & "' no tiene el método '" & eNombreMetodo & "'.")
                Else
                    Return Nothing
                End If
            End If
        End Function

        ''' <summary>
        ''' Crea una instancia de un objeto a partir del nombre de la clase
        ''' </summary>
        ''' <param name="eNombreClase">Nombre de la clase a instanciar</param>
        ''' <param name="eParametros">Si el objeto necesita parámetros en el constructor, array con los
        ''' parámetros necesarios para la invocación</param>
        ''' <returns>Nothing si no se pudo realizar la operación, o el objeto creado</returns>            
        Public Function clase2Objeto(ByVal eNombreClase As String, _
                                     Optional ByVal eParametros() As Object = Nothing) As Object
            Try
                Dim elTipo As Type = Type.GetType(eNombreClase)
                Dim elObjeto As Object = Activator.CreateInstance(elTipo, eParametros)
                Return elObjeto
            Catch ex As Exception
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Crea una instancia de un objeto a partir del nombre de la clase a instanciar
        ''' </summary>
        ''' <param name="eNombreClase">Nombre de la clase a instanciar, incluido el ensamblado que la contiene</param>
        ''' <param name="eParametros">Si el objeto necesita parámetros en el constructor, array con los
        ''' parámetros necesarios para la invocación</param>
        ''' <returns>Nothing si no se pudo realizar la operación, o el objeto creado</returns>              
        Public Function claseEnsamblado2Objeto(ByVal eNombreClase As String, _
                                               Optional ByVal eParametros() As Object = Nothing) As Object
            Dim paraDevolver As Object = Nothing

            ' Se obtiene el espacio de nombres donde está contenido la clase
            ' separando el nombre de la clase (último punto)
            Dim espacioNombres As String = eNombreClase.Substring(0, eNombreClase.LastIndexOf("."c))

            Try
                For Each unEnsamblado As Object In Assembly.GetExecutingAssembly().GetReferencedAssemblies()
                    If (unEnsamblado.Name = espacioNombres) Then
                        paraDevolver = Assembly.Load(unEnsamblado).CreateInstance(eNombreClase, eParametros)
                        Exit For
                    End If
                Next
            Catch ex As Exception
                paraDevolver = Nothing
            End Try

            Return paraDevolver
        End Function
    End Module
End Namespace
