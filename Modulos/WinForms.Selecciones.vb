Imports System.Windows.Forms
Imports ComponentFactory.Krypton.Toolkit

Namespace WinForms
    Namespace Selecciones
        Public Module modRecompilaWinFormsSelecciones

#Region " ENUMERADOS "
            ''' <summary>
            ''' Opción de selección
            ''' </summary>
            Public Enum OpcionSeleccion
                SinEscoger = -1
                Todos = 0
                Ninguno = 1
                Invertir = 2
            End Enum
#End Region

#Region " OPCIONES DE SELECCIÓN "
            ''' <summary>
            ''' Añade las opciones de selección a un objeto Combo
            ''' </summary>
            ''' <param name="eCombo">Objeto combo al que se le van a cargar las opciones de selección</param>
            Public Sub anhadirOpcionesSeleccion(ByRef eCombo As Object, _
                                                Optional ByVal eOpcionSeleccion As OpcionSeleccion = OpcionSeleccion.SinEscoger)
                With eCombo.Items
                    .Clear()
                    .Add("Todos")
                    .Add("Ninguno")
                    .Add("Invertir")
                End With
                eCombo.SelectedIndex = eOpcionSeleccion
            End Sub
#End Region

#Region " SELECCIONES UTILIZANDO EL METODO SELECT DE WINDOWS "
            ''' <summary>
            ''' Realiza la selección de los elementos de un DataGrid en función de la opción de selección
            ''' </summary>
            ''' <param name="eLista">DataGrid sobre el que se va a realizar la selección</param>
            ''' <param name="eOpcionSeleccion">Opción de selección a realizar</param>
            ''' <param name="eBarraProgreso">Barra de progreso para indicar el porcentaje de cambio</param>
            ''' <remarks></remarks>
            Public Sub marcarSeleccionadosWindowsMode(ByRef eLista As Object, _
                                                      ByVal eOpcionSeleccion As OpcionSeleccion, _
                                                      Optional ByVal eBarraProgreso As Object = Nothing)
                If TypeOf (eLista) Is DataGridView Or TypeOf (eLista) Is KryptonDataGridView Then
                    If eLista.Rows.Count > 0 Then
                        For Each elRow As DataGridViewRow In eLista.Rows
                            Select Case eOpcionSeleccion
                                Case OpcionSeleccion.Todos
                                    elRow.Selected = True

                                Case OpcionSeleccion.Ninguno
                                    elRow.Selected = False

                                Case OpcionSeleccion.Invertir
                                    elRow.Selected = Not elRow.Selected

                                Case Else

                            End Select
                        Next
                    End If
                ElseIf TypeOf (eLista) Is ListBox Or TypeOf (eLista) Is KryptonListBox Then
                    For i As Integer = 0 To CType(eLista, ListBox).Items.Count - 1
                        Select Case eOpcionSeleccion
                            Case OpcionSeleccion.Todos
                                CType(eLista, ListBox).SetSelected(i, True)

                            Case OpcionSeleccion.Ninguno
                                CType(eLista, ListBox).SetSelected(i, False)

                            Case OpcionSeleccion.Invertir
                                CType(eLista, ListBox).SetSelected(i, Not CType(eLista, ListBox).GetSelected(i))

                            Case Else

                        End Select
                    Next
                ElseIf TypeOf (eLista) Is CheckedListBox Or TypeOf (eLista) Is KryptonCheckedListBox Then
                    marcarSeleccionados(eLista, eOpcionSeleccion, eBarraProgreso)
                End If
            End Sub

            ''' <summary>
            ''' Realiza el cálculo de los elementos seleccionados en un DataGrid
            ''' </summary>
            ''' <param name="eLista">Objeto sobre el que se va a realizar la operación</param>
            ''' <returns>Número de elemento seleccionados</returns>
            Public Function calcularSeleccionadosWindowsMode(ByVal eLista As Object) As Long
                If TypeOf (eLista) Is DataGridView Or TypeOf (eLista) Is KryptonDataGridView Then
                    Return eLista.SelectedRows.Count
                ElseIf TypeOf (eLista) Is ListBox Or TypeOf (eLista) Is KryptonListBox Then
                    Return eLista.SelectedItems.Count                
                ElseIf TypeOf (eLista) Is CheckedListBox Or TypeOf (eLista) Is KryptonCheckedListBox Then
                    Return eLista.CheckedItems.Count
                Else
                    Return 0
                End If
            End Function
#End Region

#Region " SELECCIONES UTILIZANDO EL METODO CHECK "

            ''' <summary>
            ''' Marca los objetos seleccionados en un 
            ''' </summary>
            ''' <param name="eLista"></param>
            ''' <param name="eOpcionSeleccion"></param>
            ''' <param name="eBarraProgreso"></param>
            ''' <remarks></remarks>
            Public Sub marcarSeleccionados(ByRef eLista As CheckedListBox, _
                                           ByVal eOpcionSeleccion As OpcionSeleccion, _
                                           Optional ByVal eBarraProgreso As Object = Nothing)

                Cursor.Current = Cursors.WaitCursor
                Dim EstabaBarraVisible As Boolean = False

                If eBarraProgreso IsNot Nothing Then
                    EstabaBarraVisible = eBarraProgreso.Visible
                    eBarraProgreso.Visible = True
                    eBarraProgreso.Minimum = 0
                    eBarraProgreso.Maximum = eLista.Items.Count
                    eBarraProgreso.Value = 0
                End If

                ' Evitar que se repinte mientras se realizan los cambios
                eLista.SuspendLayout()

                For i As Integer = 0 To eLista.Items.Count - 1
                    Select Case eOpcionSeleccion
                        Case OpcionSeleccion.Todos
                            eLista.SetItemCheckState(i, CheckState.Checked)

                        Case OpcionSeleccion.Ninguno
                            eLista.SetItemCheckState(i, CheckState.Unchecked)

                        Case OpcionSeleccion.Invertir
                            If eLista.GetItemCheckState(i) = CheckState.Checked Then
                                eLista.SetItemCheckState(i, CheckState.Unchecked)
                            Else
                                eLista.SetItemCheckState(i, CheckState.Checked)
                            End If
                    End Select

                    If eBarraProgreso IsNot Nothing Then
                        eBarraProgreso.Value += 1
                        eBarraProgreso.refresh()
                        If CType(eBarraProgreso, Control).FindForm IsNot Nothing Then CType(eBarraProgreso, Control).FindForm.Refresh()
                        System.Windows.Forms.Application.DoEvents()
                    End If
                Next

                eLista.ResumeLayout()

                If eBarraProgreso IsNot Nothing Then
                    eBarraProgreso.Visible = EstabaBarraVisible
                    eBarraProgreso.refresh()
                    System.Windows.Forms.Application.DoEvents()
                End If

                Cursor.Current = Cursors.Default
            End Sub

            Public Sub marcarSeleccionados(ByRef eDataGrid As DataGridView, _
                                           ByVal eOpcionSeleccion As OpcionSeleccion, _
                                           ByVal eIndiceColumna As Integer, _
                                           Optional ByVal eBarraProgreso As Object = Nothing)

                Cursor.Current = Cursors.WaitCursor
                Dim EstabaBarraVisible As Boolean = False
                Dim Diccionario As New Dictionary(Of Integer, DataGridViewAutoSizeColumnMode)

                If eBarraProgreso IsNot Nothing Then
                    EstabaBarraVisible = eBarraProgreso.Visible
                    eBarraProgreso.Visible = True
                    eBarraProgreso.Minimum = 0
                    eBarraProgreso.Maximum = eDataGrid.Rows.Count + 1
                    eBarraProgreso.Value = 0
                End If

                ' MEJORA
                ' Ponemos las columnas para que no se autoredimensionen para evitar procesado extra
                For Each Columna As DataGridViewColumn In eDataGrid.Columns
                    Diccionario.Add(Columna.Index, Columna.AutoSizeMode)
                    Columna.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                Next

                ' Evito que se redibuje mientras actualizo los valores
                eDataGrid.SuspendLayout()

                For Each elRow As DataGridViewRow In eDataGrid.Rows
                    Select Case eOpcionSeleccion
                        Case 0
                            elRow.Cells(eIndiceColumna).Value = True

                        Case 1
                            elRow.Cells(eIndiceColumna).Value = False

                        Case 2
                            elRow.Cells(eIndiceColumna).Value = Not (elRow.Cells(eIndiceColumna).Value)
                    End Select

                    If eBarraProgreso IsNot Nothing Then
                        eBarraProgreso.Value += 1
                        eBarraProgreso.refresh()
                        If CType(eBarraProgreso, Control).FindForm IsNot Nothing Then CType(eBarraProgreso, Control).FindForm.Refresh()
                        System.Windows.Forms.Application.DoEvents()
                    End If
                Next

                ' Vuelvo a poner los valores por defecto
                For Each Entrada In Diccionario
                    eDataGrid.Columns(Entrada.Key).AutoSizeMode = Entrada.Value
                Next
                eDataGrid.ResumeLayout()

                If eBarraProgreso IsNot Nothing Then
                    eBarraProgreso.Visible = EstabaBarraVisible
                    eBarraProgreso.refresh()
                    System.Windows.Forms.Application.DoEvents()
                End If

                Cursor.Current = Cursors.Default
            End Sub

            Public Function calcularSeleccionados(ByVal laLista As Object, _
                                                  ByVal indiceColumna As Integer) As Long
                Dim elTotal As Integer = laLista.Rows.Count
                Dim losSeleccionados As Integer = 0
                For Each elRow As DataGridViewRow In laLista.Rows
                    If elRow.Cells(indiceColumna).Value = True Then
                        losSeleccionados += 1
                    End If
                Next
                Return losSeleccionados
            End Function
#End Region
        End Module
    End Namespace
End Namespace
