Namespace Procesos
    Namespace Ventanas
        ''' <summary>
        ''' Información que se estrae de las ventanas asociadas a los procesos
        ''' </summary>
        Public Class strInfoVentana
            Public eNombre As String = String.Empty
            Public eHWND As Long = 0

            Public eTop As Long = 0
            Public eLeft As Long = 0
            Public eBotton As Long = 0
            Public eRight As Long = 0

            Public eCMD As String = String.Empty
            Public ePID As Long = 0
        End Class
    End Namespace
End Namespace
