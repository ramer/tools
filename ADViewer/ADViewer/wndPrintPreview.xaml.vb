Imports System.Collections.ObjectModel
Imports System.Reflection
Imports System.Windows

Public Class wndPrintPreview

    Public Property objects As New clsThreadSafeObservableCollection(Of clsDirectoryObject)

    Dim fd As FlowDocument

    Private Sub btnPrint_Click(sender As Object, e As RoutedEventArgs) Handles btnPrint.Click



        'Dim printDialog As New PrintDialog
        'printDialog.PrintTicket.PageOrientation = System.Printing.PageOrientation.Landscape
        DoThePrint(fd)
        'If printDialog.ShowDialog = True Then
        '    Dim dps As IDocumentPaginatorSource = fd
        '    printDialog.PrintDocument(dps.DocumentPaginator, "ADViewer")
        'End If
    End Sub

    Private Sub wndPrintPreview_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        fd = New FlowDocument
        fd.IsColumnWidthFlexible = True

        Dim table = New Table()
        table.CellSpacing = 0
        table.BorderBrush = Brushes.Black
        table.BorderThickness = New Thickness(1)
        Dim rowGroup = New TableRowGroup()
        table.RowGroups.Add(rowGroup)
        Dim header = New TableRow()
        header.Background = Brushes.AliceBlue
        rowGroup.Rows.Add(header)

        For Each column As clsDataGridColumnInfo In preferences.Columns
            Dim tableColumn = New TableColumn()
            'configure width and such
            tableColumn.Width = New GridLength(column.Width, GridUnitType.Auto)
            table.Columns.Add(tableColumn)
            Dim hc As New Paragraph(New Run(column.Header))
            hc.FontSize = 14.0
            hc.FontFamily = New FontFamily("Segoe UI")
            hc.FontWeight = FontWeights.Bold
            Dim cell = New TableCell(hc)
            cell.BorderBrush = Brushes.Gray
            cell.BorderThickness = New Thickness(0.1)
            cell.Padding = New Thickness(5, 5, 5, 5)
            header.Cells.Add(cell)
        Next

        For Each obj In objects
            Dim tableRow = New TableRow()
            rowGroup.Rows.Add(tableRow)

            For Each column As clsDataGridColumnInfo In preferences.Columns
                Dim cell As New TableCell
                cell.BorderBrush = Brushes.Gray
                cell.BorderThickness = New Thickness(0.1)
                cell.Padding = New Thickness(5, 5, 5, 5)
                Dim first As Boolean = True
                For Each attr In column.Attributes
                    Dim t As Type = obj.GetType()
                    Dim pic() As PropertyInfo = t.GetProperties()

                    For Each pi In pic
                        If pi.Name = attr.Name Then
                            Dim value = pi.GetValue(obj)

                            If attr.Name <> "Image" Then
                                Dim p As New Paragraph(New Run(value))
                                p.FontSize = 12.0
                                p.FontFamily = New FontFamily("Segoe UI")
                                If first Then p.FontWeight = FontWeights.Bold : first = False
                                cell.Blocks.Add(p)
                            Else
                                Dim img As New Image
                                img.Source = New BitmapImage(New Uri("pack://application:,,,/" & value))
                                img.Width = 32
                                img.Height = 32
                                Dim p As New BlockUIContainer(img)
                                cell.Blocks.Add(p)
                            End If
                        End If
                    Next

                Next

                tableRow.Cells.Add(cell)
            Next
        Next

        'For Each row As DataRow In dataTable.Rows
        '    Dim tableRow = New TableRow()
        '    rowGroup.Rows.Add(tableRow)

        '    For Each column As DataColumn In dataTable.Columns
        '        Dim value = row(column).ToString()
        '        'mayby some formatting is in order
        '        Dim cell = New TableCell(New Paragraph(New Run(value)))
        '        tableRow.Cells.Add(cell)
        '    Next
        'Next


        'For Each column As clsDataGridColumnInfo In preferences.Columns
        '    Dim tableColumn = New TableColumn()
        '    'configure width and such
        '    table.Columns.Add(tableColumn)
        '    Dim hc As New Paragraph(New Run(column.Header))
        '    hc.FontWeight = FontWeights.Bold
        '    Dim cell = New TableCell(hc)
        '    header.Cells.Add(cell)
        'Next

        fd.Blocks.Add(table)


        fdsvPrintPreview.Document = fd


    End Sub
End Class
