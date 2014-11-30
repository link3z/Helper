Namespace Ficheros
    Namespace CarpetasEspeciales
        Public Module FicherosCarpetasEspeciales
#Region "ENUMERADOS"
            ''' <summary>
            ''' Enumerado con todas las carpetas especiales disponibles en el sistema
            ''' </summary>
            ''' <remarks></remarks>
            Public Enum mceIDLPaths
                CSIDL_ALTSTARTUP = &H1D ' * CSIDL_ALTSTARTUP - File system directory that corresponds to the user's nonlocalized Startup program group. (All Users\Startup?)
                CSIDL_APPDATA = &H1A ' * CSIDL_APPDATA - File system directory that serves as a common repository for application-specific data. A common path is C:\WINNT\Profiles\username\Application Data.
                CSIDL_BITBUCKET = &HA ' * CSIDL_BITBUCKET - Virtual folder containing the objects in the user's Recycle Bin.
                CSIDL_COMMON_ALTSTARTUP = &H1E ' * CSIDL_COMMON_ALTSTARTUP - File system directory that corresponds to the nonlocalized Startup program group for all users. Valid only for Windows NT systems.
                CSIDL_COMMON_APPDATA = &H23 ' * CSIDL_COMMON_APPDATA - Version 5.0. Application data for all users. A common path is C:\WINNT\Profiles\All Users\Application Data.
                CSIDL_COMMON_DESKTOPDIRECTORY = &H19 ' * CSIDL_DESKTOPDIRECTORY - File system directory used to physically store file objects on the desktop (not to be confused with the desktop folder itself). A common path is C:\WINNT\Profiles\username\Desktop
                CSIDL_COMMON_DOCUMENTS = &H2E ' * CSIDL_COMMON_DOCUMENTS - File system directory that contains documents that are common to all users. A common path is C:\WINNT\Profiles\All Users\Documents. Valid only for Windows NT systems.
                CSIDL_COMMON_FAVORITES = &H1F ' * CSIDL_COMMON_FAVORITES - File system directory that serves as a common repository for all users' favorite items. Valid only for Windows NT systems.
                CSIDL_COMMON_PROGRAMS = &H17 ' * CSIDL_COMMON_PROGRAMS - File system directory that contains the directories for the common program groups that appear on the Start menu for all users. A common path is c:\WINNT\Profiles\All Users\Start Menu\Programs. Valid only for Windows NT systems.
                CSIDL_COMMON_STARTMENU = &H16 ' * CSIDL_COMMON_STARTMENU - File system directory that contains the programs and folders that appear on the Start menu for all users. A common path is C:\WINNT\Profiles\All Users\Start Menu. Valid only for Windows NT systems.
                CSIDL_COMMON_STARTUP = &H18 ' * CSIDL_COMMON_STARTUP - File system directory that contains the programs that appear in the Startup folder for all users. A common path is C:\WINNT\Profiles\All Users\Start Menu\Programs\Startup. Valid only for Windows NT systems.
                CSIDL_COMMON_TEMPLATES = &H2D ' * CSIDL_COMMON_TEMPLATES - File system directory that contains the templates that are available to all users. A common path is C:\WINNT\Profiles\All Users\Templates. Valid only for Windows NT systems.
                CSIDL_COOKIES = &H21 ' * CSIDL_COOKIES - File system directory that serves as a common repository for Internet cookies. A common path is C:\WINNT\Profiles\username\Cookies.
                CSIDL_DESKTOPDIRECTORY = &H10 ' * CSIDL_COMMON_DESKTOPDIRECTORY - File system directory that contains files and folders that appear on the desktop for all users. A common path is C:\WINNT\Profiles\All Users\Desktop. Valid only for Windows NT systems.
                CSIDL_FAVORITES = &H6 ' * CSIDL_FAVORITES - File system directory that serves as a common repository for the user's favorite items. A common path is C:\WINNT\Profiles\username\Favorites.
                CSIDL_FONTS = &H14 ' * CSIDL_FONTS - Virtual folder containing fonts. A common path is C:\WINNT\Fonts.
                CSIDL_HISTORY = &H22 ' * CSIDL_HISTORY - File system directory that serves as a common repository for Internet history items.
                CSIDL_INTERNET_CACHE = &H20 ' * CSIDL_INTERNET_CACHE - File system directory that serves as a common repository for temporary Internet files. A common path is C:\WINNT\Profiles\username\Temporary Internet Files.
                CSIDL_LOCAL_APPDATA = &H1C ' * CSIDL_LOCAL_APPDATA - Version 5.0. File system directory that serves as a data repository for local (non-roaming) applications . A common path is C:\WINNT\Profiles\username\Local Settings\Application Data.
                CSIDL_PROGRAMS = &H2 ' * CSIDL_PROGRAMS - File system directory that contains the user's program groups (which are also file system directories). A common path is C:\WINNT\Profiles\username\Start Menu\Programs.
                CSIDL_PROGRAM_FILES = &H26 ' * CSIDL_PROGRAM_FILES - Version 5.0. Program Files folder. A common path is C:\Program Files.
                CSIDL_PROGRAM_FILES_COMMON = &H2B ' * CSIDL_PROGRAM_FILES_COMMON - Version 5.0. A folder for components that are shared across applications. A common path is C:\Program Files\Common. Valid only for Windows NT and Windows® 2000 systems.
                CSIDL_PERSONAL = &H5 ' * CSIDL_PERSONAL - File system directory that serves as a common repository for documents. A common path is C:\WINNT\Profiles\username\My Documents.
                CSIDL_RECENT = &H8 ' * CSIDL_RECENT - File system directory that contains the user's most recently used documents. A common path is C:\WINNT\Profiles\username\Recent. To create a shortcut in this folder, use SHAddToRecentDocs. In addition to creating the shortcut, this function updates the shell's list of recent documents and adds the shortcut to the Documents submenu of the Start menu.
                CSIDL_SENDTO = &H9 ' * CSIDL_SENDTO - File system directory that contains Send To menu items. A common path is c:\WINNT\Profiles\username\SendTo.
                CSIDL_STARTUP = &H7 ' * CSIDL_STARTUP - File system directory that corresponds to the user's Startup program group. The system starts these programs whenever any user logs onto Windows NT or starts Windows® 95. A common path is C:\WINNT\Profiles\username\Start Menu\Programs\Startup.
                CSIDL_STARTMENU = &HB ' * CSIDL_STARTMENU - File system directory containing Start menu items. A common path is c:\WINNT\Profiles\username\Start Menu.
                CSIDL_SYSTEM = &H25 ' * CSIDL_SYSTEM - Version 5.0. System folder. A common path is C:\WINNT\SYSTEM32.
                CSIDL_TEMPLATES = &H15 ' * CSIDL_TEMPLATES - File system directory that serves as a common repository for document templates.
                CSIDL_WINDOWS = &H24 ' * CSIDL_WINDOWS - Version 5.0. Windows directory or SYSROOT. This corresponds to the %windir% or %SYSTEMROOT% environment variables. A common path is C:\WINNT.
            End Enum
#End Region

#Region "API WINDOWS"
            ''' <summary>
            ''' Función para obtener el nombre de la carpeta especial, utilizada exclusivamente desde el módulo                
            ''' </summary>
            ''' <param name="hWnd"></param>
            ''' <param name="lpszPath"></param>
            ''' <param name="nFolder"></param>
            ''' <param name="fCreate"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Private Declare Function SHGetSpecialFolderPath Lib "SHELL32.DLL" Alias "SHGetSpecialFolderPathA" (ByVal hWnd As IntPtr, _
                                                                                                               ByVal lpszPath As String, _
                                                                                                               ByVal nFolder As Integer, _
                                                                                                               ByVal fCreate As Boolean) As Boolean
#End Region

            ''' <summary>
            ''' Función que, a partir del enumerado que especifica la carpeta que se quiere obtener, obtiene
            ''' la ruta para la Carpeta especifica para el usuario y sistema operativo
            ''' </summary>
            ''' <param name="SpecialFolder"></param>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Function ObtenerRutaCarpetaEspecial(ByVal SpecialFolder As mceIDLPaths) As String
                Dim Ret As Long
                Dim Trash As String
                Trash = Space$(260)

                Try
                    Ret = SHGetSpecialFolderPath(0, Trash, SpecialFolder, False)
                    If Trim$(Trash) <> Chr(0) Then
                        Trash = Left$(Trash, InStr(Trash, Chr(0)) - 1) & "\"
                    End If

                    Return Trash
                Catch ex As Exception
                    Throw New Exception("No se ha podido localizar la carpeta.", ex)
                End Try
            End Function
        End Module
    End Namespace
End Namespace