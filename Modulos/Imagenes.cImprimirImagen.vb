Imports System.Windows.Forms
Imports System.Drawing.Printing
Imports System.Drawing

Namespace Imagenes
    ''' <summary>
    ''' Clase para facilitar la impresión de imágenes
    ''' </summary>
    Public Class cImprimirImagen
#Region "DECLARACIONES"
        ''' <summary>
        ''' Imagen que se va a imprimir
        ''' </summary>
        Private Property ImagenAImprimir As System.Drawing.Image

        ''' <summary>
        ''' Determina si la imagen se tiene que rotar antes de imprimirse
        ''' </summary>
        Private Property RotarImagen As Boolean

        ''' <summary>
        ''' DPI utilizados para la renderización
        ''' </summary>
        Private Const RENDER_DPI As Integer = 300
#End Region

#Region "CONSTRUCTORES"
        Public Sub New(ByVal eImagen As System.Drawing.Image, _
                       ByVal eRotar As Boolean)
            ImagenAImprimir = eImagen
            RotarImagen = eRotar
        End Sub
#End Region

#Region " IMPLEMENTACION "
        Public Function ImprimirImagen() As Boolean
            ' SIGUIENTE. No redimensionar imagen si es muy pequeña (menor que folio ancho/alto depende)
            ' SIGUIENTE SIGUIENTE : Vista previa
            Dim pd As New PrintDialog()
            pd.AllowPrintToFile = False
            pd.AllowSomePages = False
            pd.PrinterSettings.FromPage = 1
            pd.PrinterSettings.ToPage = 1
            If pd.ShowDialog() = DialogResult.OK Then
                Dim printDocument As New PrintDocument()
                printDocument.PrintController = New StandardPrintController()
                printDocument.PrinterSettings = pd.PrinterSettings

                ' Margenes de impresión de imagenes (Centesimas de pulgada (1" = 2.54 cm))
                With pd.PrinterSettings.DefaultPageSettings.Margins
                    .Bottom = 100
                    .Left = 100
                    .Right = 100
                    .Top = 100
                End With

                AddHandler printDocument.PrintPage, AddressOf ManejadorImpresora

                Cursor.Current = Cursors.WaitCursor
                printDocument.Print()
                Cursor.Current = Cursors.[Default]
                Return True
            Else
                Return False
            End If
        End Function

        Private Sub AutoRotate(ByVal image As System.Drawing.Image, ByVal bounds As RectangleF)
            If (image.Height > image.Width And bounds.Width > bounds.Height) Or (image.Width > image.Height And bounds.Height > bounds.Width) Then
                image.RotateFlip(RotateFlipType.Rotate270FlipNone)
            End If
        End Sub

        Private Sub ManejadorImpresora(ByVal sender As Object, ByVal e As PrintPageEventArgs)

            If ImagenAImprimir IsNot Nothing Then
                If RotarImagen Then
                    AutoRotate(ImagenAImprimir, e.Graphics.VisibleClipBounds)
                End If

                Dim xFactor As Single = ((e.Graphics.VisibleClipBounds.Width / 100) * ImagenAImprimir.HorizontalResolution) / ImagenAImprimir.Width
                Dim yFactor As Single = ((e.Graphics.VisibleClipBounds.Height / 100) * ImagenAImprimir.VerticalResolution) / ImagenAImprimir.Height

                Dim scalePercentage As Single = If((yFactor > xFactor), xFactor, yFactor)
                Dim optimalDPI As Integer = CInt(Math.Truncate(RENDER_DPI * scalePercentage))

                e.Graphics.ScaleTransform(scalePercentage, scalePercentage)
                e.Graphics.DrawImage(ImagenAImprimir, 0, 0)

                e.HasMorePages = False
            End If
        End Sub
#End Region
    End Class

End Namespace
