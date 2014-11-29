Namespace FicheroInicio
    Public Module modRecompilaFicheroInicio

#Region " DECLARACIONES API WINDOWS "
        ''' <summary>
        ''' Para usarla en las funciones GetSection(s)
        ''' </summary>
        Private sBuffer As String

        Private Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
        Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
        Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
        Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Integer, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
        Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer
        Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Integer, ByVal lpFileName As String) As Integer
        Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Integer, ByVal lpString As Integer, ByVal lpFileName As String) As Integer
#End Region

#Region " METODOS "
        ''' <summary>
        ''' Borrar una clave o entrada de un fichero INI            
        ''' Si no se indica sKey, se borrará la sección indicada en sSection
        ''' En otro caso, se supone que es la entrada (clave) lo que se quiere borrar
        ''' Para borrar una sección se debería usar IniDeleteSection
        ''' </summary>
        ''' <param name="eIniFile">Fichero .INI</param>
        ''' <param name="eSection">Sección</param>
        ''' <param name="eKey">Clave</param>
        ''' <remarks></remarks>
        Public Sub IniDeleteKey(ByVal eIniFile As String, _
                                ByVal eSection As String, _
                                Optional ByVal eKey As String = "")
            If Len(eKey) = 0 Then
                Call WritePrivateProfileString(eSection, 0, 0, eIniFile)
            Else
                Call WritePrivateProfileString(eSection, eKey, 0, eIniFile)
            End If
        End Sub

        ''' <summary>
        ''' Borrar una sección de un fichero INI 
        ''' </summary>
        ''' <param name="eIniFile">Fichero .INI</param>
        ''' <param name="eSection">Sección</param>
        ''' <remarks></remarks>
        Public Sub IniDeleteSection(ByVal eIniFile As String, _
                                    ByVal eSection As String)
            Call WritePrivateProfileString(eSection, 0, 0, eIniFile)
        End Sub

        ''' <summary>
        ''' Devuelve el valor de una clave de un fichero INI
        ''' </summary>
        ''' <param name="eFileName">El fichero INI</param>
        ''' <param name="eSection">La sección de la que se quiere leer</param>
        ''' <param name="eKeyName">Clave</param>
        ''' <param name="eDefault">Valor opcional que devolverá si no se encuentra la clave</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function IniGet(ByVal eFileName As String, _
                               ByVal eSection As String, _
                               ByVal eKeyName As String, _
                               Optional ByVal eDefault As String = "") As String

            Dim ParaDevolver As String = eDefault

            Dim sRetVal As String = New String(Chr(0), 255)
            Dim ret As Integer = GetPrivateProfileString(eSection, eKeyName, eDefault, sRetVal, Len(sRetVal), eFileName)

            If ret <> 0 Then
                Try
                    ParaDevolver = Left(sRetVal, ret)
                Catch ex As Exception
                    ParaDevolver = eDefault
                End Try
            End If

            Return ParaDevolver
        End Function

        ''' <summary>
        ''' Guarda los datos de configuración
        ''' Los parámetros son los mismos que en LeerIni
        ''' Siendo sValue el valor a guardar
        ''' </summary>
        ''' <param name="eFileName">El fichero INI</param>
        ''' <param name="eSection">La sección de la que se quiere leer</param>
        ''' <param name="eKeyName">Clave</param>
        ''' <param name="eValue">Valor a guardar</param>
        ''' <remarks></remarks>
        Public Function IniWrite(ByVal eFileName As String, _
                                 ByVal eSection As String, _
                                 ByVal eKeyName As String, _
                                 ByVal eValue As String)
            Call WritePrivateProfileString(eSection, eKeyName, eValue, eFileName)
            Return eValue
        End Function

        ''' <summary>
        ''' Lee una sección entera de un fichero INI        
        ''' Adaptada para devolver un array de string       
        '''
        ''' Esta función devolverá un array de índice cero
        ''' con las claves y valores de la sección
        ''' </summary>
        ''' <param name="eFileName">Nombre del fichero INI</param>
        ''' <param name="eSection">Nombre de la sección a leer</param>
        ''' <returns>
        '''   Un array con el nombre de la clave y el valor
        '''   Para leer los datos:
        '''       For i = 0 To UBound(elArray) -1 Step 2
        '''           sClave = elArray(i)
        '''           sValor = elArray(i+1)
        '''       Next
        ''' </returns>
        ''' <remarks></remarks>
        Public Function IniGetSection(ByVal eFileName As String, _
                                      ByVal eSection As String) As String()
            Dim aSeccion(0) As String

            sBuffer = New String(ChrW(0), 32767)

            Dim n As Integer = GetPrivateProfileSection(eSection, sBuffer, sBuffer.Length, eFileName)
            If n > 0 Then
                ' Cortar la cadena al número de caracteres devueltos
                ' menos los dos últimos que indican el final de la cadena
                sBuffer = sBuffer.Substring(0, n - 2).TrimEnd()
                ' Cada elemento estará separado por un Chr(0)
                ' y cada valor estará en la forma: clave = valor
                aSeccion = sBuffer.Split(New Char() {ChrW(0), "="c})
            End If

            Return aSeccion
        End Function

        ''' <summary>
        ''' Devuelve todas las secciones de un fichero INI      
        ''' Adaptada para devolver un array de string        
        ''' Esta función devolverá un array con todas las secciones del fichero
        ''' </summary>
        ''' <param name="eFileName">Nombre del fichero INI</param>
        ''' <returns>
        '''  Un array con todos los nombres de las secciones
        '''   La primera sección estará en el elemento 1,
        '''   por tanto, si el array contiene cero elementos es que no hay secciones
        ''' </returns>
        ''' <remarks></remarks>
        Public Function IniGetSections(ByVal eFileName As String) As String()
            Dim aSections(0) As String

            sBuffer = New String(ChrW(0), 32767)

            ' Esta función del API no está definida en el fichero TXT
            Dim n As Integer = GetPrivateProfileSectionNames(sBuffer, sBuffer.Length, eFileName)
            If n > 0 Then
                ' Cortar la cadena al número de caracteres devueltos
                ' menos los dos últimos que indican el final de la cadena
                sBuffer = sBuffer.Substring(0, n - 2).TrimEnd()
                aSections = sBuffer.Split(ChrW(0))
            End If

            Return aSections
        End Function
#End Region
    End Module
End Namespace
