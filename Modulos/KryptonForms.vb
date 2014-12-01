Imports System.Windows.Forms
Imports ComponentFactory.Krypton.Toolkit

Namespace _KryptonForms
    Public Module modRecompilaKryptonForms
        ''' <summary>
        ''' Limpia un control Krypton
        ''' </summary>
        ''' <param name="sender">Control que se va a limpiar</param>
        Public Sub LimpiarControl(ByVal sender As System.Object)

            If TypeOf (sender) Is KryptonTextBox Then
                DirectCast(sender, KryptonTextBox).Clear()
            ElseIf TypeOf (sender) Is KryptonDateTimePicker Then
                DirectCast(sender, KryptonDateTimePicker).Value = DateTime.Now
                DirectCast(sender, KryptonDateTimePicker).ValueNullable = Nothing
            ElseIf TypeOf sender Is KryptonComboBox Then
                DirectCast(sender, KryptonComboBox).Text = ""
                DirectCast(sender, KryptonComboBox).SelectedItem = Nothing
                DirectCast(sender, KryptonComboBox).SelectedIndex = -1
            End If
        End Sub

        ''' <summary>
        ''' Realiza el efecto de esconder un panel
        ''' </summary>
        ''' <param name="eSplitter">Splitter que contiene el panel que a ocultar</param>
        ''' <param name="eHeader">Cabecera que a ocultar</param>
        ''' <param name="eInvertirPanel">Inversión del panel</param>
        Public Sub EsconderPanel(ByVal eSplitter As Object, _
                                 ByVal eHeader As Object, _
                                 Optional ByVal eInvertirPanel As Boolean = False)
            Try
                If eSplitter.Orientation = Orientation.Vertical Then
                    EsconderPanelVertical(eSplitter, eHeader, eInvertirPanel)
                Else
                    EsconderPanelHorizontal(eSplitter, eHeader, eInvertirPanel)
                End If
            Catch ex As Exception
            End Try
        End Sub

        ''' <summary>
        ''' Realiza el efecto de esconder un panel
        ''' </summary>
        ''' <param name="eSplitter">Splitter que contiene el panel que a ocultar</param>
        ''' <param name="eHeader">Cabecera que a ocultar</param>
        ''' <param name="eInvertirPanel">Inversión del panel</param>
        Private Sub EsconderPanelVertical(ByVal eSplitter As Object, _
                                          ByVal eHeader As Object, _
                                          Optional ByVal eInvertirPanel As Boolean = False)
            Try                
                eSplitter.SuspendLayout()

                Dim newWidth As Integer = eHeader.Height - eHeader.Panel.Height

                If eSplitter.tag Is Nothing Then
                    eSplitter.Panel1.Tag = eSplitter.FixedPanel
                    eSplitter.FixedPanel = FixedPanel.Panel1

                    eHeader.Tag = eHeader.Width
                    eHeader.Panel.Tag = eSplitter.IsSplitterFixed
                    eSplitter.IsSplitterFixed = True

                    With eSplitter
                        If eInvertirPanel Then
                            .FixedPanel = FixedPanel.Panel2
                            .tag = .Panel2MinSize
                            .Panel2MinSize = newWidth
                            .SplitterDistance = .Width

                            eHeader.HeaderPositionPrimary = 3
                            eHeader.ButtonSpecs(0).Edge = 0
                        Else
                            .FixedPanel = FixedPanel.Panel1
                            .tag = .Panel1MinSize
                            .Panel1MinSize = newWidth
                            .SplitterDistance = 0

                            eHeader.HeaderPositionPrimary = 3
                            eHeader.ButtonSpecs(0).Edge = 0
                        End If
                    End With

                Else
                    With eSplitter
                        .FixedPanel = .Panel1.Tag
                        .IsSplitterFixed = eHeader.Panel.Tag

                        If eInvertirPanel Then
                            .Panel2MinSize = .tag
                            .SplitterDistance = .Width - eHeader.Tag - .SplitterWidth

                            .Panel1MinSize = .tag
                            .SplitterDistance = eHeader.Tag
                        End If

                        .tag = Nothing
                    End With

                    eHeader.HeaderPositionPrimary = 0
                    eHeader.ButtonSpecs(0).Edge = 2
                End If

                eSplitter.ResumeLayout()
            Catch ex As Exception
            End Try
        End Sub

        ''' <summary>
        ''' Realiza el efecto de esconder un panel
        ''' </summary>
        ''' <param name="eSplitter">Splitter que contiene el panel que a ocultar</param>
        ''' <param name="eHeader">Cabecera que a ocultar</param>
        ''' <param name="eInvertirPanel">Inversión del panel</param>ram>
        Private Sub esconderPanelHorizontal(ByVal eSplitter As Object, _
                                            ByVal eHeader As Object, _
                                            Optional ByVal eInvertirPanel As Boolean = False)
            Try
                eSplitter.SuspendLayout()

                If eSplitter.tag Is Nothing Then
                    eSplitter.Panel1.Tag = eSplitter.FixedPanel
                    Dim newHeight As Integer = eHeader.Height - eHeader.Panel.Height

                    eHeader.Tag = eHeader.Height
                    eHeader.Panel.Tag = eSplitter.IsSplitterFixed
                    eSplitter.IsSplitterFixed = True

                    With eSplitter
                        If eInvertirPanel Then
                            .FixedPanel = FixedPanel.Panel1
                            .tag = .Panel1MinSize
                            .Panel1MinSize = newHeight
                            .SplitterDistance = 0
                        Else
                            .FixedPanel = FixedPanel.Panel2
                            .tag = .Panel2MinSize
                            .Panel2MinSize = newHeight
                            .SplitterDistance = .Height
                        End If
                    End With
                Else
                    With eSplitter
                        .FixedPanel = .Panel1.Tag
                        .IsSplitterFixed = eHeader.Panel.Tag

                        If eInvertirPanel Then
                            .Panel1MinSize = .tag
                            .SplitterDistance = eHeader.Tag
                        Else
                            .Panel2MinSize = .tag
                            .SplitterDistance = .Height - eHeader.Tag - .SplitterWidth
                        End If

                        .tag = Nothing
                    End With
                End If

                eSplitter.ResumeLayout()
            Catch ex As Exception
            End Try
        End Sub

        ''' <summary>
        ''' Solicita un dato por pantalla utilizando un KryptonInputBox
        ''' </summary>
        ''' <param name="eMensaje">Texto a mostrar para la solicitud del dato</param>
        ''' <param name="eTitulo">Título para el formulario</param>
        ''' <param name="ePredefinido">Dato predefinido</param>
        ''' <returns>Dato introducido por el usuario</returns>
        Public Function pedirTexto(ByVal eMensaje As String, _
                                   ByVal eTitulo As String, _
                                   Optional ByVal ePredefinido As String = "") As String
            Dim paraDevolver As String = KryptonInputBox.Show(eMensaje, eTitulo, ePredefinido)
            Return paraDevolver
        End Function

        ''' <summary>
        ''' Muestra un mensaje utilizando un KryptonMessageBox
        ''' </summary>
        ''' <param name="eTitulo">Título para el formulario</param>
        ''' <param name="eTexto">Texto del mensaje</param>
        ''' <param name="eBoton">Botón o botones a mostrar</param>
        ''' <param name="eIcono">Icono de la ventana</param>
        ''' <param name="ePredefinido">Boton predefinido</param>
        ''' <returns>DialogResult resultante</returns>
        Public Function MostrarMensaje(ByVal eTexto As String, _
                           ByVal eTitulo As String, _
                           ByVal eBoton As System.Windows.Forms.MessageBoxButtons, _
                           ByVal eIcono As System.Windows.Forms.MessageBoxIcon, _
                           Optional ByVal ePredefinido As System.Windows.Forms.MessageBoxDefaultButton = Nothing) As System.Windows.Forms.DialogResult
            If ePredefinido <> Nothing Then
                Return KryptonMessageBox.Show(eTexto.Trim & vbCrLf & " ", eTitulo, eBoton, eIcono, ePredefinido)
            Else
                Return KryptonMessageBox.Show(eTexto.Trim & vbCrLf & " ", eTitulo, eBoton, eIcono)

            End If
        End Function
    End Module
End Namespace